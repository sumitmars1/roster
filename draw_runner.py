from __future__ import annotations

import json
import shutil
import traceback
import uuid
from pathlib import Path
from typing import Any

from deepseek_harness import DeepSeekHarness

import book_runner
import config
import projects
import storage
from draw_update import apply_draw_patches, inspect_target_book
from limits import CONTEXT_WINDOW, MAX_OUTPUT_TOKENS

DRAW_JOBS_DIR = config.DATA / "draw_jobs"

DRAW_SYSTEM_PROMPT = """你是一名专业体育赛事秩序册抽签更新专家（触发自抽签后更新技能）。你的唯一任务是读取用户提供的抽签结果表，并对照目标秩序册中的表格与上下文，精确生成赛程签位替换补丁（submit_draw_patches）。

严格规则：
1. 先调用 mcp__draw__list_documents 并完整读取抽签结果文件中的全部工作表、页面或图片，整理出所有组别与签号（如 A1->某队名、部门A1->某队名、B2->某队名）的映射。
2. 调用 mcp__draw__inspect_target_book 查看目标秩序册的所有表格摘要与候选段落，再按 table_index 分页读取赛程、分组、对阵相关表格的完整行与上下文列（日期、场序、组别、赛别、对阵、场地）。
3. 上下文严格判定是否替换：
   - 【必须替换】：小组赛或首轮按签位对阵的占位符（如 `A1 VS A2`、`A1—A2`、`部门A1 vs 部门A2` 等），替换后保留原连接符风格（如 `动也文化 VS 钡栩传媒` 或 `动也文化——钡栩传媒`）。注意结合所在行的【组别】列区分布局（如部门A组 vs 街镇A组）。
   - 【严禁误改（原样保留）】：
     * 场序胜负链引用：`68胜——69胜`、`70负——71负`、`第1场胜—第2场胜` 等。
     * 小组出线后的名次交叉/半决赛/决赛/排位赛对阵：当【赛别】或上下文指明是小组名次出线对阵（如 `交叉赛`/`半决赛`/`排位赛`中的 `A组第一名 VS B组第二名`，或交叉轮 `A1 VS B2` 实际含义是“A组第1名对B组第2名”），绝对不能把“A组第一名”误替为1号签队伍！
4. 把识别出的应保留但未触碰的代表性胜负/名次引用写入 preserved_examples，再调用 mcp__draw__submit_draw_patches 提交逐单元格精确补丁。"""


def run_draw_update(
    target_docx: Path,
    draw_files: list[tuple[str, Path]],
    out_docx: Path,
    out_pdf: Path,
    *,
    job_id: str | None = None,
    note: str = "",
) -> dict[str, Any]:
    """Run DSH draw-update inspection & replacement on target_docx using uploaded draw_files."""
    if not config.API_KEY:
        raise RuntimeError("服务端未配置 DASHSCOPE_API_KEY，无法调用抽签更新引擎。")
    target_docx = Path(target_docx).resolve()
    if not target_docx.is_file():
        raise FileNotFoundError(f"未找到目标秩序册文件: {target_docx}")
    if not draw_files:
        raise ValueError("请至少上传一份抽签结果文件。")

    jid = job_id or f"draw_{uuid.uuid4().hex[:10]}"
    staged = [(name, path, "draw") for name, path in draw_files]
    book_runner.stage_material_items(
        jid,
        staged,
        title=f"抽签更新-{target_docx.name}",
        extra_fields={"mode": "draw"},
    )
    work_dir = config.TASKS / jid

    # Pre-inspect target .docx tables & paragraphs for MCP pagination
    inspect_data = inspect_target_book(target_docx)
    storage.write_json(work_dir / "book_inspect.json", inspect_data)

    patch_file = book_runner.build_dsh_patch(
        work_dir,
        mode="draw",
        task_id=jid,
        system_prompt=DRAW_SYSTEM_PROMPT,
        server_name="draw",
    )

    extra_note = f"\n用户补充说明：{note.strip()}\n" if note.strip() else ""
    files_list = "\n".join(f"- {name}" for name, _ in draw_files)
    prompt = f"""请更新秩序册《{target_docx.name}》中的抽签赛程占位：
抽签结果文件：
{files_list}
{extra_note}
工作步骤：
1. 调用 mcp__draw__list_documents 并读取全部抽签结果。
2. 调用 mcp__draw__inspect_target_book 获取表格列表，逐张读取相关赛程/分组表格行与上下文。
3. 核对组别、赛别上下文，构建精确的签位替换 patches 与 preserved_examples，调用 mcp__draw__submit_draw_patches 提交。"""

    harness = None
    try:
        harness = DeepSeekHarness(
            profile="sdk-minimal",
            patches=(patch_file,),
            dsh_home=str(work_dir / "dsh-home"),
            cwd=str(work_dir),
            provider="book",
            model=config.MODEL,
            reasoning_effort="off",
            max_tokens=MAX_OUTPUT_TOKENS,
            request_timeout_seconds=config.REQUEST_TIMEOUT_SECONDS,
            shutdown_timeout_seconds=3.0,
            env={"DASHSCOPE_API_KEY": config.API_KEY, "EVENT_BOOK_DATA_DIR": str(config.DATA)},
        )
        harness.start()
        session = harness.start_session(f"draw-{jid}")
        result = session.run(prompt)
        harness.close()
        harness = None
    finally:
        if harness is not None:
            try:
                harness.close()
            except Exception:
                pass

    patches_payload = storage.read_json(work_dir / "draw_patches.json")
    if not isinstance(patches_payload, dict):
        tail = (getattr(result, "final_response", "") or "")[:400] if result else ""
        raise RuntimeError(f"Agent 未提交抽签替换补丁（draw_patches.json）。{tail}")

    applied = apply_draw_patches(
        target_docx,
        patches_payload.get("patches") or [],
        out_docx,
        out_pdf,
        preserved_examples=patches_payload.get("preserved_examples") or [],
    )
    applied["agent_summary"] = patches_payload.get("summary") or ""
    return applied


def run_project_draw_update(
    project_id: str,
    draw_files: list[tuple[str, Path]],
    *,
    source_version_id: str | None = None,
    note: str = "",
) -> dict[str, Any]:
    """Apply post-draw update to an existing project's generated order book and record a new version."""
    project = projects.get_project(project_id)
    if not project:
        raise ValueError(f"项目不存在: {project_id}")
    src_ver = projects.get_version(project, source_version_id)
    if not src_ver:
        raise ValueError("该项目尚无可更新的已生成秩序册版本，请先生成秩序册。")
    src_docx = projects.project_dir(project_id) / src_ver["docx_rel"]
    if not src_docx.is_file():
        raise FileNotFoundError(f"未找到原秩序册 Word 文件: {src_docx}")

    vid = projects.allocate_version_id(project_id)
    vdir = projects.project_version_dir(project_id, vid)
    projects.update_project(project_id, {"status": "running", "stage": "正在核对抽签结果并更新已生成秩序册", "error": ""})
    projects.append_project_trace(project_id, {"stage": "draw_update", "version_id": vid, "from": src_ver["id"]})

    # Also archive uploaded draw result files inside version dir
    draw_dir = vdir / "draw_sources"
    draw_dir.mkdir(parents=True, exist_ok=True)
    archived: list[tuple[str, Path]] = []
    for idx, (orig_name, p) in enumerate(draw_files, 1):
        dest = draw_dir / f"{idx:02d}_{Path(orig_name).name}"
        shutil.copy2(p, dest)
        archived.append((orig_name, dest))

    try:
        out_docx = vdir / "book.docx"
        out_pdf = vdir / "book.pdf"
        res = run_draw_update(
            src_docx,
            archived,
            out_docx,
            out_pdf,
            job_id=f"draw_{project_id}_{vid}",
            note=note,
        )
        import pymupdf

        with pymupdf.open(out_pdf) as pdf_f:
            pdf_pages = pdf_f.page_count

        version_info = {
            "id": vid,
            "kind": "draw_update",
            "parent_version_id": src_ver["id"],
            "created_at": storage.now(),
            "event_title": src_ver.get("event_title") or project["name"],
            "template_id": src_ver.get("template_id") or project["template_id"],
            "template_name": src_ver.get("template_name") or "",
            "docx_rel": f"versions/{vid}/book.docx",
            "pdf_rel": f"versions/{vid}/book.pdf",
            "pdf_pages": pdf_pages,
            "toc": res.get("toc") or src_ver.get("toc") or [],
            "chapters": src_ver.get("chapters") or [],
            "missing_items": src_ver.get("missing_items") or [],
            "excluded_items": src_ver.get("excluded_items") or [],
            "unrecognized_files": src_ver.get("unrecognized_files") or [],
            "draw_summary": {
                "replaced_count": res["replaced_count"],
                "skipped_count": res["skipped_count"],
                "preserved_count": res["preserved_count"],
                "changes": res["changes"],
                "preserved": res["preserved"],
                "agent_summary": res["agent_summary"],
            },
        }
        projects.record_version(project_id, version_info)
        projects.append_project_trace(
            project_id,
            {"stage": "draw_done", "version_id": vid, "replaced": res["replaced_count"]},
        )
        return version_info
    except Exception as exc:
        err = f"{type(exc).__name__}: {exc}"
        (vdir / "draw-error.log").write_text(traceback.format_exc(), encoding="utf-8")
        projects.update_project(
            project_id,
            {
                "status": "ready",
                "stage": f"抽签更新失败：{err[:160]}",
                "error": err[:500],
            },
        )
        raise


def create_standalone_draw_job(
    target_docx_src: Path,
    target_filename: str,
    draw_files: list[tuple[str, Path]],
    *,
    note: str = "",
) -> dict[str, Any]:
    """Create and execute a standalone post-draw update job for an uploaded .docx file."""
    DRAW_JOBS_DIR.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex[:12]
    jdir = DRAW_JOBS_DIR / job_id
    jdir.mkdir(parents=True, exist_ok=True)

    target_copy = jdir / f"original_{Path(target_filename).name}"
    shutil.copy2(target_docx_src, target_copy)

    staged_draws: list[tuple[str, Path]] = []
    for idx, (orig, p) in enumerate(draw_files, 1):
        d = jdir / f"draw_{idx:02d}_{Path(orig).name}"
        shutil.copy2(p, d)
        staged_draws.append((orig, d))

    out_docx = jdir / "updated_book.docx"
    out_pdf = jdir / "updated_book.pdf"
    res = run_draw_update(target_copy, staged_draws, out_docx, out_pdf, job_id=f"drawjob_{job_id}", note=note)
    meta = {
        "id": job_id,
        "target_name": Path(target_filename).name,
        "created_at": storage.now(),
        "docx_name": out_docx.name,
        "pdf_name": out_pdf.name,
        **res,
    }
    storage.write_json(jdir / "job.json", meta)
    return meta
