from __future__ import annotations

import copy
import json
import re
import shutil
import sys
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import openpyxl
import pymupdf
from deepseek_harness import DeepSeekHarness
from docx import Document
from docx.oxml.ns import qn
from PIL import Image

import book_assembler
import config
import documents
import harness_runner
import projects
import renderer
import storage
import templates_store
from limits import CONTEXT_WINDOW, MAX_OUTPUT_TOKENS
from presets import PRESETS

BOOK_MCP_SERVER = Path(__file__).resolve().with_name("book_mcp.py")

PLAN_SYSTEM_PROMPT = """你是一名专业体育赛事秩序册整册编排总控专家。你的唯一目标是基于用户上传的全部比赛原始资料和所选模板，快速研判整批文件语义，提交结构严谨的整册编排规划（submit_book_plan）。

必须遵循以下准则：
1. 第一步先调用 mcp__book__inspect_template 和 mcp__book__list_documents。提示词中已预提取全部资料的工作表名、行数、Word 纲要与首页摘录，优先直接结合该摘要规划。
2. 对每份非图片文件只需做最低限度的结构核实即可：
   - 图片封面/底图：用 read_document(mode="image") 查看视觉内容；若封面图片已经自带赛事主标题和“秩序册”字样，cover.mode 设为 "image"；否则设 "background" 交由系统合成封面。仅当用户单独提供了浅色/半透明内页衬底图时才填入 watermark.file_id，严禁把深色封面图同时设为内页水印。
   - Word（.docx/.doc）：优先调用 inspect_docx_outline（联合文档可用 start_pattern / end_pattern 精确切分章节），严禁对 Word 文件调用 mode="image" 或 mode="layout"。
   - Excel/PDF：每份文档最多用 read_document(unit=1, limit=25) 抽样核实口径 1 次即可。**绝对禁止在规划阶段逐行翻页通读报名表长名单！** 完整名单抽取会在下一阶段并发分片执行。
3. 严禁只凭文件名猜测归属！交叉核对【竞赛规程】【分组情况】【竞赛日程表】与【报名汇总表】：
   - 例如报名表中有多张工作表（如“街镇”“部门”“部门女子”“Sheet1”），只有被规程、分组或赛程实际纳入比赛的工作表才列入 roster_scope.sources[].included_sheets，并填好 group_map（工作表名精确对应竞赛组别名）；
   - 未实际参赛或空白工作表必须放入 excluded_sheets，并在 excluded_items 中说明依据。
4. 对可成型的 Word 资料与 Excel 赛程/分组表，直接按文件映射承接，绝不改动赛程中未抽签的签位占位（如 A1 VS A2）。
5. 严格参照模板 chapters 顺序输出规划：缺少资料的可选/必填章节写入 missing_items；确实无法辨认用途的文件写入 unrecognized_files。核实完毕立即调用 submit_book_plan。"""


def build_dsh_patch(
    work_dir: Path,
    *,
    mode: str,
    task_id: str,
    system_prompt: str,
    server_name: str,
) -> str:
    patch = [
        {"id": name, "disabled": True}
        for name in (
            "llm-deepseek",
            "tool-fs",
            "tool-working-directory",
            "terminal-bash",
            "terminal-pwsh",
            "persistent-bash",
            "persistent-pwsh",
        )
    ]
    patch += [
        {
            "id": "system-prompt",
            "config": {
                "includeHarnessIdentity": False,
                "includeRuntimeContext": False,
                "personaPrefix": system_prompt,
            },
        },
        {
            "insert": [
                {"id": "skill", "name": "@deepseek-ai/dsh-skill"},
                {
                    "id": "skill-filesystem",
                    "name": "@deepseek-ai/dsh-skill-filesystem",
                    "config": {
                        "includeDefaultRoots": False,
                        "watch": False,
                        "customSkillDirs": [str(config.SKILLS), str(config.ROOT / "skills")],
                    },
                },
                {"id": "tool-skill", "name": "@deepseek-ai/dsh-tool-skill"},
            ]
        },
        {
            "insert": [
                {
                    "id": f"{server_name}-mcp",
                    "name": "@deepseek-ai/dsh-mcp-client",
                    "config": {
                        "serverName": server_name,
                        "transport": "stdio",
                        "command": sys.executable,
                        "args": [str(BOOK_MCP_SERVER), "--task", task_id, "--mode", mode],
                        "cwd": str(work_dir),
                        "toolCallTimeoutMs": 120_000,
                        "startupTimeoutMs": 30_000,
                        "failOnStartupError": True,
                    },
                }
            ]
        },
        {
            "insert": [
                {
                    "id": "llm-pi-ai",
                    "name": "@deepseek-ai/dsh-llm-pi-ai",
                    "config": {
                        "providers": {
                            "book": {
                                "api": "openai-responses",
                                "baseURL": config.BASE_URL,
                                "retryPolicy": {
                                    "mode": "normal",
                                    "maxRetries": 2,
                                    "retryableCodes": [
                                        "EMPTY_RESPONSE",
                                        "RATE_LIMIT",
                                        "SERVER",
                                        "TIMEOUT",
                                        "TRANSPORT",
                                        "MALFORMED_RESPONSE",
                                    ],
                                },
                                "apiKeyEnv": "DASHSCOPE_API_KEY",
                                "models": [
                                    {
                                        "id": config.MODEL,
                                        "input": ["text", "image"],
                                        "contextWindow": CONTEXT_WINDOW,
                                        "maxTokens": MAX_OUTPUT_TOKENS,
                                        "reasoningEfforts": {"off": "none", "low": "low"},
                                    }
                                ],
                            }
                        }
                    },
                }
            ]
        },
    ]
    patch_path = work_dir / f"harness-{mode}.patch.json"
    patch_path.write_text(json.dumps(patch, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(patch_path)


def stage_material_items(
    task_id: str,
    materials: list[tuple[str, Path, str]],
    *,
    preset: str = "inline-roster",
    title: str = "",
    group: str = "",
    extra_fields: dict[str, Any] | None = None,
) -> dict[str, Path]:
    """Create a DSH-compatible internal task directory containing copied files & inventories."""
    tdir = config.TASKS / task_id
    if tdir.exists():
        shutil.rmtree(tdir)
    tdir.mkdir(parents=True, exist_ok=True)

    saved: list[dict[str, Any]] = []
    fid_to_path: dict[str, Path] = {}
    for index, (orig_name, src_path, role_hint) in enumerate(materials, 1):
        suffix = Path(orig_name).suffix.lower() or src_path.suffix.lower()
        stored_name = f"f{index:02d}{suffix}"
        dest = tdir / stored_name
        shutil.copy2(src_path, dest)
        fid = f"F{index}"
        fid_to_path[fid] = dest
        try:
            units = documents.inventory(dest, bulk=True)
        except Exception as exc:
            units = [{"unit": "raw", "name": orig_name, "error": str(exc)}]
        saved.append(
            {
                "id": fid,
                "name": orig_name,
                "stored_name": stored_name,
                "suffix": suffix,
                "size": dest.stat().st_size,
                "status": "waiting",
                "role_hint": role_hint,
                "units": units,
            }
        )

    state: dict[str, Any] = {
        "id": task_id,
        "internal": True,
        "title": title or "秩序册项目",
        "group": group or "",
        "preset": preset if preset in PRESETS else "inline-roster",
        "preset_name": PRESETS.get(preset, ("",))[0],
        "status": "queued",
        "files": saved,
        "session_id": f"task-{task_id}",
        "usage": {"input_tokens": 0, "output_tokens": 0},
        "result": None,
        "review": None,
        "error": None,
        "answers": [],
        "active_file": None,
        "reads": {},
        "created_at": storage.now(),
        "updated_at": storage.now(),
        "started_at": None,
        "finished_at": None,
    }
    if extra_fields:
        state.update(extra_fields)
    storage.write_json(tdir / "state.json", state)
    return fid_to_path


def _digest_docx_summary(path: Path) -> list[str]:
    lines: list[str] = []
    if path.suffix.lower() == ".doc":
        with documents.converted_doc(path) as conv:
            return _digest_docx_summary(conv)
    doc = Document(path)
    p_count = 0
    t_count = 0
    for child in doc.element.body:
        if child.tag == qn("w:p"):
            txt = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            has_img = bool(child.findall(f".//{qn('w:drawing')}") or child.findall(f".//{qn('w:pict')}"))
            if (txt or has_img) and len(lines) < 16:
                tag = "[含图]" if has_img else ""
                lines.append(f"    段落p{p_count}{tag}: {txt[:80]}")
            p_count += 1
        elif child.tag == qn("w:tbl"):
            rows = child.findall(qn("w:tr"))
            hdr = []
            if rows:
                hdr = [
                    "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()[:16]
                    for tc in rows[0].findall(qn("w:tc"))[:8]
                ]
            if len(lines) < 22:
                lines.append(f"    表格t{t_count}（{len(rows)}行）表头: {' | '.join(hdr)}")
            t_count += 1
    lines.append(f"    （共 {p_count} 段、{t_count} 张表格）")
    return lines


def _digest_one_file(file_meta: dict[str, Any], path: Path) -> str:
    fid = str(file_meta.get("id") or "")
    name = str(file_meta.get("name") or path.name)
    hint = str(file_meta.get("role_hint") or "auto")
    hint_note = f"（用户提示角色：{hint}）" if hint != "auto" else ""
    header = f"- {fid}: {name}{hint_note}"
    suffix = path.suffix.lower()
    details: list[str] = []
    try:
        if suffix in {".xlsx", ".xls", ".et"}:
            units = documents.inventory(path, bulk=True)
            for u in units[:10]:
                rows, _, _, _ = documents.extract(path, u, 1, fid, bulk=True, row_limit=5)
                details.append(f"    工作表「{u['label']}」（{u['size']}行）:")
                for _, vals in rows:
                    row_txt = " | ".join(str(v) for v in vals if str(v).strip())
                    if row_txt:
                        details.append(f"      {row_txt[:110]}")
            if len(units) > 10:
                details.append(f"    ...其余 {len(units) - 10} 张工作表")
        elif suffix in {".docx", ".doc"}:
            details.extend(_digest_docx_summary(path))
        elif suffix == ".pdf":
            with pymupdf.open(path) as pdf:
                total_pages = len(pdf)
                sample_lines: list[str] = []
                for pno in range(min(2, total_pages)):
                    page = pdf[pno]
                    txt_lines = [ln.strip() for ln in page.get_text("text", sort=True).splitlines() if ln.strip()]
                    img_flag = " [含大图/扫描]" if documents.page_has_large_image(page) else ""
                    details.append(f"    第 {pno + 1}/{total_pages} 页{img_flag}: {' / '.join(txt_lines[:6])[:120]}")
                    sample_lines.extend(txt_lines)
        elif suffix in {".png", ".jpg", ".jpeg", ".webp"}:
            with Image.open(path) as im:
                details.append(f"    图片尺寸 {im.width}x{im.height}（需用 read_document(mode='image') 查看）")
        else:
            units = file_meta.get("units") or []
            details.append(f"    {len(units)} 个单元")
    except Exception as exc:
        details.append(f"    [摘要跳过: {str(exc)[:80]}]")
    return "\n".join([header, *details]) if details else header


def _build_stage1_digest(work_dir: Path, files_meta: list[dict[str, Any]]) -> str:
    if not files_meta:
        return ""
    max_workers = min(8, max(1, len(files_meta)))
    ordered: list[str] = [""] * len(files_meta)
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futs = {
            pool.submit(_digest_one_file, f, work_dir / str(f.get("stored_name") or f["name"])): idx
            for idx, f in enumerate(files_meta)
        }
        for fut in as_completed(futs):
            ordered[futs[fut]] = fut.result()
    return "\n".join(line for line in ordered if line)


def _plan_prompt(project: dict[str, Any], template: dict[str, Any], extra_note: str = "") -> str:
    planner_task_id = f"bookplan_{project['id']}_{project['_pending_vid']}"
    state = storage.read_task(planner_task_id) or {}
    work_dir = config.TASKS / planner_task_id
    files_meta = state.get("files") or []
    files_str = _build_stage1_digest(work_dir, files_meta) or "\n".join(
        f"- {f['id']}: {f['name']}" for f in files_meta
    )
    chap_lines = [
        f"- {c.get('id')}: {c.get('title')} (kind={c.get('kind')}, required={bool(c.get('required'))}"
        + (", split_by_group=true" if c.get("split_by_group_in_toc") else "")
        + ")"
        for c in (template.get("chapters") or [])
    ]
    chap_block = "\n".join(chap_lines)
    note_block = f"\n用户补充说明：{extra_note}\n" if extra_note.strip() else ""
    return f"""请为以下赛事项目快速制定整册秩序册编排规划：
项目/赛事参考名称：{project.get('name')}
选定模板：{template.get('name')}（ID={template.get('id')}，默认名单母版={template.get('roster_preset')}）
模板章节清单：
{chap_block}
{note_block}
本次上传的全部原始资料与预提取结构摘要：
{files_str}

快速核实与提交流程：
1. 先调用 mcp__book__inspect_template 与 mcp__book__list_documents。
2. 图片用 read_document(mode="image") 查看；Word（.docx/.doc）用 inspect_docx_outline（禁止调用 image/layout）；其余尚未读过的非图片文档最多调用 1 次 read_document(unit=1, limit=25) 核实表头与内容（严禁逐行通读长表格）。
3. 结合规程与分组确定 roster_scope.sources（含 included_sheets、excluded_sheets、精确的 group_map）及各章节文件映射后，立即调用 mcp__book__submit_book_plan 提交最终规划。"""


def _run_stage1_planner(
    project: dict[str, Any],
    template: dict[str, Any],
    planner_task_id: str,
    extra_note: str = "",
) -> dict[str, Any]:
    if not config.API_KEY:
        raise RuntimeError("服务端未配置 DASHSCOPE_API_KEY，无法调用智能编排引擎。")
    work_dir = config.TASKS / planner_task_id
    patch_file = build_dsh_patch(
        work_dir,
        mode="plan",
        task_id=planner_task_id,
        system_prompt=PLAN_SYSTEM_PROMPT,
        server_name="book",
    )
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
        session = harness.start_session(f"plan-{planner_task_id}")
        result = session.run(_plan_prompt(project, template, extra_note))
        harness.close()
        harness = None
    finally:
        if harness is not None:
            try:
                harness.close()
            except Exception:
                pass

    plan = storage.read_json(work_dir / "book_plan.json")
    if not isinstance(plan, dict):
        tail = (getattr(result, "final_response", "") or "")[:400] if result else ""
        raise RuntimeError(f"Agent 未提交整册规划书（book_plan.json）。{tail}")
    return plan


def _filter_scoped_units(units: list[dict[str, Any]], included_sheets: list[str] | None) -> list[dict[str, Any]]:
    if not included_sheets:
        return units
    allowed = {str(s).strip() for s in included_sheets if str(s).strip()}
    filtered = [
        u
        for u in units
        if str(u.get("label") or u.get("name") or u.get("unit") or "").strip() in allowed
    ]
    return filtered or units


_EXCEL_HEADER_KEYWORDS = (
    "序号", "单位", "姓名", "性别", "身份证", "号码", "球衣", "代表队", "队伍", "组别", "电话", "年龄",
)
_BANNER_KEYWORDS = ("组", "男", "女", "甲", "乙", "丙", "代表", "队")


def _cell_is_empty(val: Any) -> bool:
    return val is None or str(val).strip() == ""


def _load_excel_sheets(
    src_path: Path,
) -> list[tuple[str, list[list[tuple[Any, str | None]]], list[tuple[int, int, int, int]]]]:
    suffix = src_path.suffix.lower()
    if suffix == ".et":
        with documents.converted_doc(src_path) as conv:
            return _load_excel_sheets(conv)
    out: list[tuple[str, list[list[tuple[Any, str | None]]], list[tuple[int, int, int, int]]]] = []
    if suffix == ".xls":
        book = documents.legacy_workbook(src_path)
        for sh in book.sheets():
            grid = [
                [(documents.legacy_cell_text(book, sh.cell(r, c)), "@") for c in range(sh.ncols)]
                for r in range(sh.nrows)
            ]
            while grid and all(_cell_is_empty(v) for v, _ in grid[-1]):
                grid.pop()
            max_r = len(grid)
            merges = [
                (rlo + 1, min(rhi, max_r), clo + 1, chi)
                for rlo, rhi, clo, chi in getattr(sh, "merged_cells", ())
                if rlo + 1 <= max_r
            ]
            out.append((str(sh.name), grid, merges))
        return out
    if suffix == ".xlsx":
        wb = openpyxl.load_workbook(src_path, data_only=True, keep_links=False)
        try:
            for ws in wb.worksheets:
                max_r = ws.max_row or 0
                max_c = ws.max_column or 0
                grid = [
                    [(cell.value, cell.number_format) for cell in row]
                    for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c)
                ]
                while grid and all(_cell_is_empty(v) for v, _ in grid[-1]):
                    grid.pop()
                trimmed_r = len(grid)
                merges = [
                    (rng.min_row, min(rng.max_row, trimmed_r), rng.min_col, rng.max_col)
                    for rng in ws.merged_cells.ranges
                    if rng.min_row <= trimmed_r
                ]
                out.append((str(ws.title), grid, merges))
        finally:
            wb.close()
        return out
    return out


def _find_header_end(
    grid: list[list[tuple[Any, str | None]]],
    merges: list[tuple[int, int, int, int]],
) -> int:
    total = len(grid)
    if total <= 1:
        return total
    header_end = 1
    found_keyword = False
    for r_idx in range(1, min(5, total) + 1):
        vals = [str(v).strip() for v, _ in grid[r_idx - 1] if not _cell_is_empty(v)]
        joined = " ".join(vals)
        if len(vals) >= 3 and any(kw in joined for kw in _EXCEL_HEADER_KEYWORDS):
            header_end = r_idx
            found_keyword = True
            break
    if not found_keyword:
        for r_idx in range(1, min(3, total) + 1):
            vals = [str(v).strip() for v, _ in grid[r_idx - 1] if not _cell_is_empty(v)]
            if len(vals) >= 3:
                header_end = r_idx
                break
    for r1, r2, _, _ in merges:
        if r1 <= header_end < r2 and r2 <= 6:
            header_end = min(total, r2)
    return header_end


def _split_sheet_ranges(
    grid: list[list[tuple[Any, str | None]]],
    merges: list[tuple[int, int, int, int]],
    header_end: int,
    *,
    target_rows: int = 50,
    max_rows: int = 62,
) -> list[tuple[int, int, str]]:
    total = len(grid)
    data_start = header_end + 1
    if total < data_start:
        return []
    if total - header_end <= max_rows:
        return [(data_start, total, "")]

    forbidden_starts: set[int] = set()
    merge_starts: set[int] = set()
    for r1, r2, c1, _ in merges:
        if r2 > r1:
            for rr in range(r1 + 1, r2 + 1):
                forbidden_starts.add(rr)
            if c1 <= 3:
                merge_starts.add(r1)

    banner_at: dict[int, str] = {}
    active_banner_before: dict[int, str] = {}
    cur_banner = ""
    for r in range(data_start, total + 1):
        active_banner_before[r] = cur_banner
        vals = [str(v).strip() for v, _ in grid[r - 1] if not _cell_is_empty(v)]
        if 1 <= len(vals) <= 2 and any(kw in "".join(vals) for kw in _BANNER_KEYWORDS):
            text = " ".join(vals)[:60]
            banner_at[r] = text
            cur_banner = text

    def _row_pair(r_num: int) -> tuple[str, str]:
        row = grid[r_num - 1]
        c0 = str(row[0][0]).strip() if len(row) > 0 and not _cell_is_empty(row[0][0]) else ""
        c1 = str(row[1][0]).strip() if len(row) > 1 and not _cell_is_empty(row[1][0]) else ""
        return c0, c1

    def _is_blank_row(r_num: int) -> bool:
        return all(_cell_is_empty(v) for v, _ in grid[r_num - 1])

    ranges: list[tuple[int, int, str]] = []
    cur = data_start
    while cur <= total:
        if total - cur + 1 <= max_rows:
            ranges.append((cur, total, active_banner_before.get(cur, "")))
            break
        lo = min(total + 1, cur + 32)
        hi = min(total, cur + max_rows)
        best_cand: int | None = None
        best_key: tuple[int, int] | None = None
        for cand in range(lo, hi + 1):
            if cand in forbidden_starts:
                continue
            if cand in banner_at:
                score = 4
            elif (cand - 1 >= data_start and _is_blank_row(cand - 1)) or cand in merge_starts:
                score = 3
            else:
                prev_pair = _row_pair(cand - 1) if cand - 1 >= data_start else ("", "")
                cur_pair = _row_pair(cand)
                changed = any(
                    cur_pair[i] and cur_pair[i] != prev_pair[i] for i in (0, 1)
                )
                score = 2 if changed else 1
            rank = (score, -abs(cand - (cur + target_rows)))
            if best_key is None or rank > best_key:
                best_key = rank
                best_cand = cand
        if best_cand is None:
            for cand in range(min(total, cur + 31), cur + 9, -1):
                if cand not in forbidden_starts:
                    best_cand = cand
                    break
        if best_cand is None:
            for cand in range(cur + max_rows + 1, total + 1):
                if cand not in forbidden_starts:
                    best_cand = cand
                    break
        if best_cand is None:
            best_cand = total + 1
        ranges.append((cur, best_cand - 1, active_banner_before.get(cur, "")))
        cur = best_cand
    return ranges


def _write_xlsx_shard(
    out_path: Path,
    sheet_title: str,
    grid: list[list[tuple[Any, str | None]]],
    merges: list[tuple[int, int, int, int]],
    header_end: int,
    r_start: int,
    r_end: int,
) -> None:
    wb = openpyxl.Workbook()
    try:
        ws = wb.active
        ws.title = (sheet_title or "Sheet1")[:31]

        def _write_row(src_row_idx: int, dst_row_idx: int) -> None:
            for c_idx, (val, fmt) in enumerate(grid[src_row_idx - 1], 1):
                cell = ws.cell(row=dst_row_idx, column=c_idx, value=val)
                if isinstance(val, str) and not fmt:
                    cell.number_format = "@"
                elif fmt:
                    cell.number_format = fmt

        if r_start <= header_end:
            for r in range(1, len(grid) + 1):
                _write_row(r, r)
            for r1, r2, c1, c2 in merges:
                if r2 > r1 or c2 > c1:
                    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        else:
            for r in range(1, header_end + 1):
                _write_row(r, r)
            offset = (header_end + 1) - r_start
            for r in range(r_start, r_end + 1):
                _write_row(r, r + offset)
            for r1, r2, c1, c2 in merges:
                if (r2 > r1 or c2 > c1) and r1 >= 1 and r2 <= header_end:
                    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
                elif (r2 > r1 or c2 > c1) and r1 >= r_start and r2 <= r_end:
                    ws.merge_cells(
                        start_row=r1 + offset,
                        start_column=c1,
                        end_row=r2 + offset,
                        end_column=c2,
                    )
        wb.save(out_path)
    finally:
        wb.close()


def _resolve_sheet_group(sheet_name: str, norm_gmap: dict[str, str], fallback_group: str) -> str:
    clean = sheet_name.strip()
    if clean in norm_gmap:
        return norm_gmap[clean]
    if norm_gmap:
        compact = re.sub(r"\s+", "", clean)
        hits = [
            v
            for k, v in norm_gmap.items()
            if re.sub(r"\s+", "", k) == compact or clean in k or k in clean
        ]
        unique_hits = list(dict.fromkeys(hits))
        if len(unique_hits) == 1:
            return unique_hits[0]
        if len(norm_gmap) == 1:
            return next(iter(norm_gmap.values()))
    return fallback_group


def _prepare_excel_shards(
    src_path: Path,
    orig_name: str,
    *,
    included_sheets: list[str],
    excluded_sheets: list[str],
    norm_gmap: dict[str, str],
    fallback_group: str,
    out_dir: Path,
    file_prefix: str,
) -> list[dict[str, Any]] | None:
    try:
        sheets = _load_excel_sheets(src_path)
    except Exception:
        return None
    if not sheets:
        return None

    inc_set = {str(s).strip() for s in included_sheets if str(s).strip()}
    exc_set = {str(s).strip() for s in excluded_sheets if str(s).strip()}
    if inc_set:
        selected = [(idx, name, grid, merges) for idx, (name, grid, merges) in enumerate(sheets, 1) if name.strip() in inc_set]
    else:
        selected = [(idx, name, grid, merges) for idx, (name, grid, merges) in enumerate(sheets, 1) if name.strip() not in exc_set]
    if not selected:
        selected = [(idx, name, grid, merges) for idx, (name, grid, merges) in enumerate(sheets, 1)]

    non_empty = [item for item in selected if item[2]]
    if non_empty:
        selected = non_empty

    prepared: list[dict[str, Any]] = []
    for s_idx, s_title, grid, merges in selected:
        g_name = _resolve_sheet_group(s_title, norm_gmap, fallback_group)
        if not grid:
            out_path = out_dir / f"{file_prefix}_s{s_idx}.xlsx"
            _write_xlsx_shard(out_path, s_title, grid, merges, 0, 1, 0)
            prepared.append(
                {
                    "path": out_path,
                    "display_name": f"{Path(orig_name).stem}__{s_title}.xlsx",
                    "group": g_name,
                    "shard_id": out_path.stem,
                    "context_hint": f"工作表「{s_title}」（所属组别：{g_name or '按规程'}）",
                    "units": documents.inventory(out_path, bulk=True),
                }
            )
            continue

        header_end = _find_header_end(grid, merges)
        ranges = _split_sheet_ranges(grid, merges, header_end)
        if not ranges:
            ranges = [(1, len(grid), "")]
            eff_header = 0
        else:
            eff_header = header_end

        multi = len(ranges) > 1
        for c_idx, (r_start, r_end, banner) in enumerate(ranges, 1):
            suffix_tag = f"_p{c_idx}" if multi else ""
            out_path = out_dir / f"{file_prefix}_s{s_idx}{suffix_tag}.xlsx"
            _write_xlsx_shard(out_path, s_title, grid, merges, eff_header, r_start, r_end)
            part_label = f"第 {r_start}–{r_end} 行（已附第 1–{eff_header} 行表头，原表共 {len(grid)} 行）" if eff_header and multi else f"工作表「{s_title}」共 {len(grid)} 行"
            hint = f"{part_label}；所属组别：{g_name or '按规程'}"
            if banner:
                hint += f"；前序分组标题：{banner}（仅作上下文，勿编造未出现的人员）"
            prepared.append(
                {
                    "path": out_path,
                    "display_name": f"{Path(orig_name).stem}__{s_title}{suffix_tag}.xlsx",
                    "group": g_name,
                    "shard_id": out_path.stem,
                    "context_hint": hint,
                    "units": documents.inventory(out_path, bulk=True),
                }
            )
    return prepared


def _prepare_pdf_shards(
    src_path: Path,
    orig_name: str,
    group_name: str,
    out_dir: Path,
    file_prefix: str,
    *,
    max_pages: int = 10,
    chunk_pages: int = 8,
) -> list[dict[str, Any]] | None:
    try:
        with pymupdf.open(src_path) as doc:
            total = len(doc)
            if total <= max_pages:
                return None
            prepared: list[dict[str, Any]] = []
            c_idx = 0
            for start_p in range(0, total, chunk_pages):
                c_idx += 1
                end_p = min(total - 1, start_p + chunk_pages - 1)
                out_path = out_dir / f"{file_prefix}_p{c_idx}.pdf"
                chunk = pymupdf.open()
                try:
                    chunk.insert_pdf(doc, from_page=start_p, to_page=end_p)
                    chunk.save(out_path)
                finally:
                    chunk.close()
                prepared.append(
                    {
                        "path": out_path,
                        "display_name": f"{Path(orig_name).stem}_p{c_idx}.pdf",
                        "group": group_name,
                        "shard_id": out_path.stem,
                        "context_hint": f"PDF 第 {start_p + 1}–{end_p + 1} 页（共 {total} 页）；所属组别：{group_name or '按规程'}",
                        "units": documents.inventory(out_path, bulk=True),
                    }
                )
            return prepared
    except Exception:
        return None


def _run_stage2_group_task(
    group_task_id: str,
    group_name: str,
    group_items: list[dict[str, Any]],
    *,
    event_title: str,
    preset: str,
    instructions: str,
) -> tuple[str, dict[str, Any] | None, Path | None]:
    tdir = config.TASKS / group_task_id
    if tdir.exists():
        shutil.rmtree(tdir)
    tdir.mkdir(parents=True, exist_ok=True)

    saved_files: list[dict[str, Any]] = []
    for f_idx, item in enumerate(group_items, 1):
        src = Path(item["path"])
        suffix = src.suffix.lower() or Path(item["display_name"]).suffix.lower()
        stored_name = f"f{f_idx:02d}{suffix}"
        dest = tdir / stored_name
        shutil.copy2(src, dest)
        units = item.get("units") or documents.inventory(dest, bulk=True)
        saved_files.append(
            {
                "id": f"F{f_idx}",
                "name": item["display_name"],
                "stored_name": stored_name,
                "suffix": suffix,
                "size": dest.stat().st_size,
                "status": "waiting",
                "units": units,
                "shard_id": item.get("shard_id") or f"{group_task_id}_F{f_idx}",
                "context_hint": item.get("context_hint") or "",
            }
        )

    state = {
        "id": group_task_id,
        "internal": True,
        "title": event_title,
        "group": (group_name or "参赛队伍")[:80],
        "instructions": instructions[:400],
        "preset": preset,
        "preset_name": PRESETS[preset][0],
        "status": "queued",
        "files": saved_files,
        "session_id": f"task-{group_task_id}",
        "usage": {"input_tokens": 0, "output_tokens": 0},
        "result": None,
        "review": None,
        "error": None,
        "answers": [],
        "active_file": None,
        "reads": {},
        "created_at": storage.now(),
        "updated_at": storage.now(),
        "started_at": None,
        "finished_at": None,
    }
    storage.write_json(tdir / "state.json", state)

    harness_runner.run_task(group_task_id)
    final_state = storage.read_task(group_task_id) or {}
    if final_state.get("status") == "review":
        q = (final_state.get("review") or {}).get("question") or "报名表存在需要人工确认的内容"
        raise RuntimeError(f"名单抽取暂停需要人工核对：{q}")
    if final_state.get("status") != "done":
        err = final_state.get("error") or "未知错误"
        raise RuntimeError(f"名单抽取阶段失败：{err}")

    roster_json = storage.read_json(tdir / "roster.json")
    full_docx = tdir / str((final_state.get("result") or {}).get("word") or "roster.docx")
    return group_name, roster_json if isinstance(roster_json, dict) else None, full_docx if full_docx.is_file() else None


def _run_stage2_rosters(
    project: dict[str, Any],
    template: dict[str, Any],
    plan: dict[str, Any],
    fid_to_path: dict[str, Path],
    fid_to_name: dict[str, str],
    roster_task_id: str,
) -> list[dict[str, Any]]:
    """Run parallel sharded roster extraction grouped cleanly by competition group."""
    scope = plan.get("roster_scope") or {}
    sources = scope.get("sources") or []
    if not sources:
        return []

    preset = str(template.get("roster_preset") or "inline-roster")
    if preset not in PRESETS:
        preset = "inline-roster"

    tdir = config.TASKS / roster_task_id
    if tdir.exists():
        shutil.rmtree(tdir)
    tdir.mkdir(parents=True, exist_ok=True)
    shard_stage_dir = tdir / "_shards"
    shard_stage_dir.mkdir(parents=True, exist_ok=True)

    planned_groups = [str(g).strip() for g in (scope.get("groups") or []) if str(g).strip()]
    single_default_group = planned_groups[0] if len(planned_groups) == 1 else ("、".join(planned_groups) or "参赛队伍")

    prepared_items: list[dict[str, Any]] = []
    scope_notes: list[str] = []
    for idx, src_item in enumerate(sources, 1):
        fid = str(src_item.get("file_id") or "")
        src_path = fid_to_path.get(fid)
        if not src_path or not src_path.is_file():
            continue
        orig_name = fid_to_name.get(fid, src_path.name)
        suffix = Path(orig_name).suffix.lower() or src_path.suffix.lower()
        inc = [str(s).strip() for s in (src_item.get("included_sheets") or []) if str(s).strip()]
        exc = [str(s).strip() for s in (src_item.get("excluded_sheets") or []) if str(s).strip()]
        norm_gmap = {
            str(k).strip(): str(v).strip()
            for k, v in (src_item.get("group_map") or {}).items()
            if str(k).strip() and str(v).strip()
        }
        if norm_gmap:
            mapping_str = "，".join(f"表[{k}]->组别[{v}]" for k, v in norm_gmap.items())
            scope_notes.append(f"{orig_name}({mapping_str})")
        if exc:
            scope_notes.append(f"排除工作表:{','.join(exc)}")

        fallback_g = next(iter(norm_gmap.values())) if len(norm_gmap) == 1 else single_default_group

        if suffix in {".xlsx", ".xls", ".et"}:
            shards = _prepare_excel_shards(
                src_path,
                orig_name,
                included_sheets=inc,
                excluded_sheets=exc,
                norm_gmap=norm_gmap,
                fallback_group=fallback_g,
                out_dir=shard_stage_dir,
                file_prefix=f"src{idx}",
            )
            if shards is not None:
                prepared_items.extend(shards)
                continue

        if suffix == ".pdf":
            pdf_shards = _prepare_pdf_shards(
                src_path,
                orig_name,
                fallback_g,
                shard_stage_dir,
                f"src{idx}",
            )
            if pdf_shards is not None:
                prepared_items.extend(pdf_shards)
                continue

        units = documents.inventory(src_path, bulk=True)
        scoped_units = _filter_scoped_units(units, inc)
        prepared_items.append(
            {
                "path": src_path,
                "display_name": orig_name,
                "group": fallback_g,
                "shard_id": f"src{idx}",
                "context_hint": "",
                "units": scoped_units,
            }
        )

    if not prepared_items:
        return []

    buckets: dict[str, list[dict[str, Any]]] = {}
    for item in prepared_items:
        g_key = str(item.get("group") or "").strip() or single_default_group
        buckets.setdefault(g_key, []).append(item)

    ordered_keys = [g for g in planned_groups if g in buckets]
    for g in buckets:
        if g not in ordered_keys:
            ordered_keys.append(g)

    instr_parts: list[str] = list(scope_notes)
    extra_instr = str(scope.get("instruction") or "").strip()
    if extra_instr:
        instr_parts.append(extra_instr)
    target_instr = "；".join(instr_parts)[:400]
    event_title = str(plan.get("event_title") or project.get("name") or "赛事名单")

    group_jobs = [
        (
            f"{roster_task_id}_g{g_idx}" if len(ordered_keys) > 1 else roster_task_id,
            g_name,
            buckets[g_name],
        )
        for g_idx, g_name in enumerate(ordered_keys, 1)
    ]

    group_outputs: list[tuple[str, dict[str, Any] | None, Path | None]] = []
    if len(group_jobs) == 1:
        tid, g_name, g_items = group_jobs[0]
        group_outputs.append(
            _run_stage2_group_task(
                tid,
                g_name,
                g_items,
                event_title=event_title,
                preset=preset,
                instructions=target_instr,
            )
        )
    else:
        ordered_results: list[tuple[str, dict[str, Any] | None, Path | None] | None] = [None] * len(group_jobs)
        with ThreadPoolExecutor(max_workers=min(4, len(group_jobs))) as pool:
            futs = {
                pool.submit(
                    _run_stage2_group_task,
                    tid,
                    g_name,
                    g_items,
                    event_title=event_title,
                    preset=preset,
                    instructions=target_instr,
                ): pos
                for pos, (tid, g_name, g_items) in enumerate(group_jobs)
            }
            for fut in as_completed(futs):
                ordered_results[futs[fut]] = fut.result()
        group_outputs = [r for r in ordered_results if r is not None]

    valid_outputs = [(g, r, d) for g, r, d in group_outputs if d is not None]
    if not valid_outputs:
        return []

    tpl_roster_chap = next(
        (c for c in (template.get("chapters") or []) if c.get("kind") == "roster"), {}
    )
    want_split = bool(tpl_roster_chap.get("split_by_group_in_toc", False))

    if len(valid_outputs) >= 2:
        if want_split:
            return [{"group": g_name, "docx": g_docx} for g_name, _, g_docx in valid_outputs if g_docx]
        if preset == "two-column-numbered":
            combined_groups: list[dict[str, Any]] = []
            for g_name, r_json, _ in valid_outputs:
                for g_obj in (r_json or {}).get("groups") or []:
                    if isinstance(g_obj, dict):
                        copy_obj = copy.deepcopy(g_obj)
                        if not copy_obj.get("title"):
                            copy_obj["title"] = g_name
                        combined_groups.append(copy_obj)
            merged_roster: dict[str, Any] = {"groups": combined_groups}
        else:
            combined_sections: list[dict[str, Any]] = []
            for _, r_json, _ in valid_outputs:
                for sec in (r_json or {}).get("sections") or []:
                    if isinstance(sec, dict):
                        combined_sections.append(copy.deepcopy(sec))
            merged_roster = {"sections": combined_sections}
        storage.write_json(tdir / "roster.json", merged_roster)
        res = renderer.render(
            roster_task_id,
            tdir,
            preset,
            merged_roster,
            event_title,
            "、".join(ordered_keys) or single_default_group,
        )
        merged_docx = Path(res["directory"]) / res["word"] if "directory" in res else (tdir / res["word"])
        return [{"group": "", "docx": merged_docx}]

    _, roster_json, full_docx = valid_outputs[0]
    if not full_docx:
        return []

    if want_split and isinstance(roster_json, dict):
        split_jobs: list[tuple[str, dict[str, Any]]] = []
        top_groups = roster_json.get("groups")
        sections = roster_json.get("sections")
        if preset == "two-column-numbered" and isinstance(top_groups, list):
            merged_two: dict[str, dict[str, Any]] = {}
            for g_obj in top_groups:
                if not isinstance(g_obj, dict):
                    continue
                g_name = str(g_obj.get("title") or g_obj.get("group") or g_obj.get("name") or "").strip() or "代表队"
                bucket = merged_two.setdefault(g_name, {**copy.deepcopy(g_obj), "teams": []})
                bucket["teams"].extend(copy.deepcopy(g_obj.get("teams") or []))
            if len(merged_two) >= 2:
                split_jobs = [(g_name, {"groups": [obj]}) for g_name, obj in merged_two.items()]
        elif isinstance(sections, list) and sections:
            has_nested = any(isinstance((s.get("data") or {}).get("groups"), list) for s in sections if isinstance(s, dict))
            if has_nested:
                base_sec = copy.deepcopy(next((s for s in sections if isinstance(s, dict)), {}))
                base_data = {k: v for k, v in (base_sec.get("data") or {}).items() if k != "groups"}
                merged_nested: dict[str, list[dict[str, Any]]] = {}
                for sec in sections:
                    if not isinstance(sec, dict):
                        continue
                    for g_obj in (sec.get("data") or {}).get("groups") or []:
                        if not isinstance(g_obj, dict):
                            continue
                        g_name = str(g_obj.get("group") or g_obj.get("title") or g_obj.get("name") or "").strip() or "各代表队"
                        merged_nested.setdefault(g_name, []).extend(copy.deepcopy(g_obj.get("teams") or []))
                if len(merged_nested) >= 2:
                    for g_name, teams in merged_nested.items():
                        sec_copy = copy.deepcopy(base_sec)
                        sec_copy["data"] = {**copy.deepcopy(base_data), "groups": [{"group": g_name, "teams": teams}]}
                        split_jobs.append((g_name, {"sections": [sec_copy]}))
            else:
                merged_flat: dict[str, list[dict[str, Any]]] = {}
                unnamed = False
                for sec in sections:
                    if not isinstance(sec, dict):
                        continue
                    data = sec.get("data") or {}
                    g_name = str(data.get("subtitle") or data.get("title") or "").strip()
                    if not g_name:
                        unnamed = True
                        break
                    merged_flat.setdefault(g_name, []).append(copy.deepcopy(sec))
                if not unnamed and len(merged_flat) >= 2:
                    for g_name, secs in merged_flat.items():
                        split_jobs.append((g_name, {"sections": secs}))

        if split_jobs:
            split_sections: list[dict[str, Any]] = []
            for g_idx, (g_name, sub_roster) in enumerate(split_jobs, 1):
                sub_dir = tdir / f"group_{g_idx}"
                sub_dir.mkdir(parents=True, exist_ok=True)
                res = renderer.render(
                    f"{roster_task_id}_g{g_idx}",
                    sub_dir,
                    preset,
                    sub_roster,
                    str(plan.get("event_title") or project.get("name") or ""),
                    g_name,
                )
                split_sections.append(
                    {
                        "group": g_name,
                        "docx": Path(res["directory"]) / res["word"] if "directory" in res else (sub_dir / res["word"]),
                    }
                )
            return split_sections

    return [{"group": "", "docx": full_docx}]


def run_book_generation(project_id: str, extra_note: str = "") -> dict[str, Any]:
    """Execute full end-to-end event book generation for a project."""
    project = projects.get_project(project_id)
    if not project:
        raise ValueError(f"项目不存在: {project_id}")
    template = templates_store.get_template(project["template_id"])
    if not template:
        raise ValueError(f"选定模板不存在: {project['template_id']}")
    materials = project.get("materials") or []
    if not materials:
        raise ValueError("请先上传赛事原始资料后再生成秩序册。")

    vid = projects.allocate_version_id(project_id)
    vdir = projects.project_version_dir(project_id, vid)
    projects.update_project(project_id, {"status": "running", "stage": "第一阶段：整体研判资料结构与范围", "error": ""})
    projects.append_project_trace(project_id, {"stage": "plan", "version_id": vid, "materials": len(materials)})

    planner_task_id = f"bookplan_{project_id}_{vid}"
    roster_task_id = f"bookroster_{project_id}_{vid}"

    try:
        items = [
            (m["name"], projects.material_file_path(project, m), m.get("role_hint") or "auto")
            for m in materials
        ]
        fid_to_path = stage_material_items(
            planner_task_id,
            items,
            preset=template.get("roster_preset") or "inline-roster",
            title=project["name"],
            extra_fields={"template_spec": template, "project_id": project_id},
        )
        fid_to_name = {f"F{idx}": m["name"] for idx, m in enumerate(materials, 1)}

        proj_ctx = {**project, "_pending_vid": vid}
        plan = _run_stage1_planner(proj_ctx, template, planner_task_id, extra_note=extra_note)
        projects.append_project_trace(
            project_id,
            {
                "stage": "plan_done",
                "event_title": plan.get("event_title"),
                "chapters": len(plan.get("chapters") or []),
                "excluded": plan.get("excluded_items") or [],
                "missing": plan.get("missing_items") or [],
            },
        )

        # Stage 2: Extract rosters via DSH + roster_skill
        projects.update_project(project_id, {"stage": "第二阶段：核对范围并整理各代表队名单"})
        roster_sections = _run_stage2_rosters(
            project, template, plan, fid_to_path, fid_to_name, roster_task_id
        )
        projects.append_project_trace(
            project_id, {"stage": "rosters_done", "groups": len(roster_sections)}
        )

        # Stage 3: Build concrete BookBuildPlan with resolved local file paths
        projects.update_project(project_id, {"stage": "第三阶段：按模板装订整册、更新目录页码并导出 Word/PDF"})
        resolved_chapters: list[dict[str, Any]] = []
        base_roster_title = next(
            (c.get("title") for c in (plan.get("chapters") or []) if c.get("kind") == "roster"),
            "各代表队名单",
        )
        roster_inserted = False

        for ch in plan.get("chapters") or []:
            kind = str(ch.get("kind") or "richtext")
            if kind in {"cover", "toc", "back_cover"}:
                continue
            if kind == "roster":
                if not roster_sections or roster_inserted:
                    continue
                roster_inserted = True
                for rs in roster_sections:
                    g = rs["group"]
                    ch_title = f"{re.sub(r'^[一二三四五六七八九十]+[、.．]', '', base_roster_title)}（{g}）" if g else base_roster_title
                    resolved_chapters.append(
                        {
                            "id": f"roster_{g or 'all'}",
                            "title": ch_title,
                            "kind": "roster",
                            "toc_visible": True,
                            "landscape": False,
                            "sources": [{"type": "docx", "path": str(rs["docx"])}],
                        }
                    )
                continue

            if kind == "notes":
                resolved_chapters.append({"id": ch.get("id") or "notes", "title": ch.get("title") or "记事", "kind": "notes", "toc_visible": False})
                continue

            resolved_sources: list[dict[str, Any]] = []
            for src in ch.get("sources") or []:
                fid = str(src.get("file_id") or "")
                p = fid_to_path.get(fid)
                if not p or not p.is_file():
                    continue
                ext = p.suffix.lower().lstrip(".")
                if ext in {"png", "jpg", "jpeg", "webp"}:
                    stype = "image"
                else:
                    stype = ext
                resolved_sources.append(
                    {
                        "type": stype,
                        "path": str(p),
                        "sheets": src.get("sheets"),
                        "start_pattern": src.get("start_pattern"),
                        "end_pattern": src.get("end_pattern"),
                        "caption": src.get("caption"),
                    }
                )
            if resolved_sources:
                resolved_chapters.append(
                    {
                        "id": ch.get("id") or kind,
                        "title": ch["title"],
                        "kind": kind,
                        "toc_visible": bool(ch.get("toc_visible", True)),
                        "landscape": ch.get("landscape", False),
                        "sources": resolved_sources,
                    }
                )

        if roster_sections and not roster_inserted:
            for rs in roster_sections:
                g = rs["group"]
                ch_title = f"各代表队名单（{g}）" if g else "各代表队名单"
                resolved_chapters.append(
                    {
                        "id": f"roster_{g or 'all'}",
                        "title": ch_title,
                        "kind": "roster",
                        "toc_visible": True,
                        "landscape": False,
                        "sources": [{"type": "docx", "path": str(rs["docx"])}],
                    }
                )

        cover_spec = plan.get("cover") or {}
        cover_fid = cover_spec.get("file_id")
        back_spec = plan.get("back_cover") or {}
        back_fid = back_spec.get("file_id")
        wm_spec = plan.get("watermark") or {}
        wm_fid = wm_spec.get("file_id")

        assembler_plan = {
            "event_title": plan.get("event_title") or project["name"],
            "subtitle": plan.get("subtitle") or "秩  序  册",
            "footer_lines": plan.get("footer_lines") or [],
            "cover": {
                "mode": cover_spec.get("mode") or ("image" if cover_fid else "generated"),
                "path": str(fid_to_path[cover_fid]) if cover_fid and cover_fid in fid_to_path else None,
            },
            "back_cover": {
                "path": str(fid_to_path[back_fid]) if back_fid and back_fid in fid_to_path else None,
            },
            "chapters": resolved_chapters,
        }

        out_docx = vdir / "book.docx"
        out_pdf = vdir / "book.pdf"
        wm_override = fid_to_path[wm_fid] if wm_fid and wm_fid in fid_to_path else None
        if wm_override and wm_fid == cover_fid and cover_spec.get("mode") == "background":
            try:
                from PIL import Image, ImageStat

                with Image.open(wm_override) as im_wm:
                    mean_luma = ImageStat.Stat(im_wm.convert("L")).mean[0]
                if mean_luma < 215.0:
                    wm_override = None
            except Exception:
                wm_override = None
        assembled = book_assembler.assemble_book(
            assembler_plan,
            template,
            out_docx,
            out_pdf,
            watermark_override=wm_override,
        )

        # Save build plan for transparency/audit
        (vdir / "book_plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

        missing_items = list(plan.get("missing_items") or [])
        unrecognized = list(plan.get("unrecognized_files") or [])
        excluded = list(plan.get("excluded_items") or [])

        version_info = {
            "id": vid,
            "kind": "full_book",
            "created_at": storage.now(),
            "event_title": assembler_plan["event_title"],
            "template_id": template["id"],
            "template_name": template["name"],
            "docx_rel": f"versions/{vid}/book.docx",
            "pdf_rel": f"versions/{vid}/book.pdf",
            "pdf_pages": assembled["pdf_pages"],
            "toc": assembled["toc"],
            "chapters": assembled["chapters"],
            "missing_items": missing_items,
            "excluded_items": excluded,
            "unrecognized_files": unrecognized,
        }
        projects.record_version(project_id, version_info)
        projects.append_project_trace(project_id, {"stage": "done", "version_id": vid, "pages": assembled["pdf_pages"]})
        return version_info
    except Exception as exc:
        err_msg = f"{type(exc).__name__}: {exc}"
        projects.append_project_trace(project_id, {"stage": "error", "error": err_msg[:500]})
        (vdir / "book-error.log").write_text(traceback.format_exc(), encoding="utf-8")
        has_prev = bool(project.get("latest_version_id"))
        projects.update_project(
            project_id,
            {
                "status": "ready" if has_prev else "error",
                "stage": f"本次重新生成失败：{err_msg[:180]}" if has_prev else "生成失败",
                "error": err_msg[:500],
            },
        )
        raise
