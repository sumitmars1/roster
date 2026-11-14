from __future__ import annotations

import shutil
import tempfile
import threading
import traceback
import uuid
from pathlib import Path
from typing import Any

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from starlette.concurrency import run_in_threadpool

import book_runner
import config
import documents
import draw_runner
import projects
import storage
import template_parser
import templates_store
from cover_art import generate_cover_from_background

router = APIRouter(prefix="/api", tags=["event-book"])

BOOK_MATERIAL_EXTENSIONS = documents.EXTENSIONS | {".docx", ".doc", ".wps", ".xlsx", ".xls", ".et", ".pdf", ".png", ".jpg", ".jpeg", ".webp"}


def _safe_filename(name: str, fallback: str) -> str:
    base = Path(name or fallback).name.strip()
    return base.replace("/", "_").replace("\\", "_") or fallback


async def _save_upload_to(upload: UploadFile, dest: Path) -> int:
    dest.parent.mkdir(parents=True, exist_ok=True)
    size = 0
    with dest.open("wb") as handle:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > config.MAX_FILE_BYTES:
                raise HTTPException(422, f"{upload.filename} 超过单份文件大小限制。")
            handle.write(chunk)
    if size <= 0:
        raise HTTPException(422, f"{upload.filename or '上传文件'} 是空文件。")
    return size


def _project_or_404(project_id: str) -> dict[str, Any]:
    proj = projects.get_project(project_id)
    if proj is None:
        raise HTTPException(404, "项目不存在。")
    return proj


# ---------------------------------------------------------------------------
# Templates & Cover Art
# ---------------------------------------------------------------------------


@router.get("/templates")
def api_list_templates():
    return {"templates": templates_store.list_templates()}


@router.get("/templates/{template_id}")
def api_get_template(template_id: str):
    tpl = templates_store.get_template(template_id)
    if tpl is None:
        raise HTTPException(404, "模板不存在。")
    return tpl


@router.get("/templates/{template_id}/watermark")
def api_template_watermark(template_id: str):
    tpl = templates_store.get_template(template_id)
    if tpl is None:
        raise HTTPException(404, "模板不存在。")
    wm_name = (tpl.get("formatting") or {}).get("watermark", {}).get("asset")
    asset = templates_store.template_asset_path(template_id, wm_name)
    if asset is None:
        raise HTTPException(404, "该模板不包含底图资产。")
    return FileResponse(asset)


@router.post("/templates/parse", status_code=201)
async def api_parse_template(
    file: UploadFile = File(...),
    name: str = Form(""),
    use_llm: bool = Form(True),
):
    fname = _safe_filename(file.filename or "", "样例秩序册.docx")
    suffix = Path(fname).suffix.lower()
    if suffix not in {".docx", ".doc", ".wps"}:
        raise HTTPException(422, "仅支持上传 Word 秩序册样例（.docx / .doc / .wps）。")

    with tempfile.TemporaryDirectory(prefix="tpl-up-") as tmp_dir:
        tmp_path = Path(tmp_dir) / fname
        try:
            await _save_upload_to(file, tmp_path)
        finally:
            await file.close()
        try:
            saved = await run_in_threadpool(
                template_parser.parse_sample_book_to_template,
                tmp_path,
                name=name.strip() or None,
                use_llm=use_llm,
            )
        except Exception as exc:
            raise HTTPException(422, f"解析模板失败：{exc}") from exc
    return saved


@router.delete("/templates/{template_id}")
def api_delete_template(template_id: str):
    ok = templates_store.delete_custom_template(template_id)
    if not ok:
        raise HTTPException(400, "无法删除内置模板或模板不存在。")
    return {"ok": True}


@router.post("/cover/preview")
async def api_generate_cover_preview(
    event_title: str = Form(...),
    subtitle: str = Form("秩  序  册"),
    footer: str = Form(""),
    file: UploadFile | None = File(None),
):
    out_dir = config.DATA / "cover_previews"
    out_dir.mkdir(parents=True, exist_ok=True)
    cid = uuid.uuid4().hex[:10]
    out_png = out_dir / f"cover_{cid}.png"
    with tempfile.TemporaryDirectory(prefix="cov-up-") as tmp_dir:
        bg_path: Path | None = None
        if file is not None and file.filename:
            fname = _safe_filename(file.filename, "bg.png")
            bg_path = Path(tmp_dir) / fname
            try:
                await _save_upload_to(file, bg_path)
            finally:
                await file.close()
        footer_lines = [ln.strip() for ln in footer.splitlines() if ln.strip()]
        await run_in_threadpool(
            generate_cover_from_background,
            bg_path,
            out_png,
            event_title.strip(),
            subtitle=subtitle.strip() or "秩  序  册",
            footer_lines=footer_lines,
        )
    return {
        "id": cid,
        "preview_url": f"/api/cover/preview/{cid}",
    }


@router.get("/cover/preview/{cover_id}")
def api_get_cover_preview(cover_id: str):
    path = config.DATA / "cover_previews" / f"cover_{cover_id}.png"
    if not path.is_file():
        raise HTTPException(404, "封面预览图不存在。")
    return FileResponse(path, media_type="image/png", filename=f"cover_{cover_id}.png")


# ---------------------------------------------------------------------------
# Projects & Materials & Full-Book Generation
# ---------------------------------------------------------------------------


@router.get("/projects")
def api_list_projects():
    return {"projects": projects.list_projects()}


@router.post("/projects", status_code=201)
def api_create_project(
    name: str = Form(...),
    template_id: str = Form(...),
):
    clean_name = name.strip()
    if not clean_name:
        raise HTTPException(422, "请填写项目或赛事名称。")
    tpl = templates_store.get_template(template_id.strip())
    if tpl is None:
        raise HTTPException(422, f"所选模板不存在：{template_id}")
    return projects.create_project(clean_name, tpl["id"])


@router.get("/projects/{project_id}")
def api_get_project(project_id: str):
    proj = _project_or_404(project_id)
    tpl = templates_store.get_template(proj["template_id"])
    return {
        **proj,
        "template": tpl,
        "trace": projects.read_project_trace(project_id, limit=80),
    }


@router.patch("/projects/{project_id}")
def api_update_project(
    project_id: str,
    name: str = Form(""),
    template_id: str = Form(""),
):
    proj = _project_or_404(project_id)
    if proj.get("status") == "running":
        raise HTTPException(409, "项目正在生成中，请稍后再修改基础设置。")
    updates: dict[str, Any] = {}
    if name.strip():
        updates["name"] = name.strip()
    if template_id.strip():
        tpl = templates_store.get_template(template_id.strip())
        if tpl is None:
            raise HTTPException(422, f"所选模板不存在：{template_id}")
        updates["template_id"] = tpl["id"]
    if not updates:
        return proj
    return projects.update_project(project_id, updates)


@router.post("/projects/{project_id}/materials", status_code=201)
async def api_upload_materials(
    project_id: str,
    files: list[UploadFile] = File(...),
    role_hint: str = Form("auto"),
    replace_material_id: str = Form(""),
):
    _project_or_404(project_id)
    if not files:
        raise HTTPException(422, "请至少选择一份资料文件。")
    if replace_material_id and len(files) > 1:
        raise HTTPException(422, "替换单份资料时一次只能上传 1 个文件。")

    added: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="mat-up-") as tmp_dir:
        tmp = Path(tmp_dir)
        for idx, up in enumerate(files, 1):
            fname = _safe_filename(up.filename or "", f"资料_{idx}")
            suffix = Path(fname).suffix.lower()
            if suffix not in BOOK_MATERIAL_EXTENSIONS:
                raise HTTPException(422, f"{fname} 的格式不受支持。")
            dest = tmp / fname
            try:
                await _save_upload_to(up, dest)
            finally:
                await up.close()
            entry = projects.add_material_file(
                project_id,
                dest,
                fname,
                role_hint=role_hint.strip() or "auto",
                replace_material_id=replace_material_id.strip() or None,
            )
            added.append(entry)
    return {"added": added, "project": projects.get_project(project_id)}


@router.delete("/projects/{project_id}/materials/{material_id}")
def api_delete_material(project_id: str, material_id: str):
    _project_or_404(project_id)
    ok = projects.remove_material(project_id, material_id)
    if not ok:
        raise HTTPException(404, "资料文件不存在。")
    return {"ok": True, "project": projects.get_project(project_id)}


@router.get("/projects/{project_id}/materials/{material_id}/raw")
def api_download_material(project_id: str, material_id: str):
    proj = _project_or_404(project_id)
    mat = next((m for m in (proj.get("materials") or []) if m.get("id") == material_id), None)
    if not mat:
        raise HTTPException(404, "资料文件不存在。")
    path = projects.material_file_path(proj, mat)
    if not path.is_file():
        raise HTTPException(404, "文件已丢失。")
    return FileResponse(path, filename=mat["name"])


def _run_generate_bg(project_id: str, extra_note: str) -> None:
    try:
        book_runner.run_book_generation(project_id, extra_note=extra_note)
    except Exception:
        pass


@router.post("/projects/{project_id}/generate", status_code=202)
def api_generate_book(
    project_id: str,
    note: str = Form(""),
):
    proj = _project_or_404(project_id)
    if proj.get("status") == "running":
        raise HTTPException(409, "当前正在生成秩序册，请等待本轮完成。")
    if not proj.get("materials"):
        raise HTTPException(422, "请先上传原始资料后再开始生成。")

    projects.update_project(
        project_id,
        {"status": "running", "stage": "已排队，准备研判全部上传资料", "error": ""},
    )
    threading.Thread(target=_run_generate_bg, args=(project_id, note.strip()), daemon=True).start()
    return projects.get_project(project_id)


def _version_file_or_404(project_id: str, version_id: str, kind: str) -> tuple[dict[str, Any], Path]:
    proj = _project_or_404(project_id)
    ver = projects.get_version(proj, None if version_id == "latest" else version_id)
    if not ver:
        raise HTTPException(404, "尚无已生成的秩序册版本。")
    rel = ver.get("docx_rel" if kind == "docx" else "pdf_rel") or ""
    path = projects.project_dir(project_id) / rel
    if not path.is_file():
        raise HTTPException(404, f"版本 {ver['id']} 的 {kind.upper()} 文件不存在。")
    return ver, path


@router.get("/projects/{project_id}/versions/{version_id}/download/docx")
def api_download_project_docx(project_id: str, version_id: str = "latest"):
    ver, path = _version_file_or_404(project_id, version_id, "docx")
    title = ver.get("event_title") or "赛事秩序册"
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{title}-秩序册-{ver['id']}.docx",
    )


@router.get("/projects/{project_id}/versions/{version_id}/download/pdf")
def api_download_project_pdf(project_id: str, version_id: str = "latest"):
    ver, path = _version_file_or_404(project_id, version_id, "pdf")
    title = ver.get("event_title") or "赛事秩序册"
    return FileResponse(
        path,
        media_type="application/pdf",
        filename=f"{title}-秩序册-{ver['id']}.pdf",
    )


@router.get("/projects/{project_id}/versions/{version_id}/preview/pdf")
def api_preview_project_pdf(project_id: str, version_id: str = "latest"):
    _ver, path = _version_file_or_404(project_id, version_id, "pdf")
    return FileResponse(
        path,
        media_type="application/pdf",
        content_disposition_type="inline",
    )


# ---------------------------------------------------------------------------
# Post-Draw Schedule Update (Project + Standalone)
# ---------------------------------------------------------------------------


def _run_project_draw_bg(project_id: str, draw_dir: Path, archived: list[tuple[str, Path]], source_vid: str | None, note: str) -> None:
    try:
        draw_runner.run_project_draw_update(
            project_id,
            archived,
            source_version_id=source_vid,
            note=note,
        )
    except Exception:
        pass
    finally:
        shutil.rmtree(draw_dir, ignore_errors=True)


@router.post("/projects/{project_id}/draw-update", status_code=202)
async def api_project_draw_update(
    project_id: str,
    files: list[UploadFile] = File(...),
    source_version_id: str = Form(""),
    note: str = Form(""),
):
    proj = _project_or_404(project_id)
    if proj.get("status") == "running":
        raise HTTPException(409, "当前项目正在执行生成或更新任务。")
    src_ver = projects.get_version(proj, source_version_id.strip() or None)
    if not src_ver:
        raise HTTPException(409, "请先生成一份秩序册，再执行抽签后更新。")
    if not files:
        raise HTTPException(422, "请上传抽签结果表文件。")

    tmp_draw_dir = Path(tempfile.mkdtemp(prefix="pdraw-"))
    archived: list[tuple[str, Path]] = []
    for idx, up in enumerate(files, 1):
        fname = _safe_filename(up.filename or "", f"抽签结果_{idx}")
        suffix = Path(fname).suffix.lower()
        if suffix not in BOOK_MATERIAL_EXTENSIONS:
            shutil.rmtree(tmp_draw_dir, ignore_errors=True)
            raise HTTPException(422, f"{fname} 的格式不受支持。")
        dest = tmp_draw_dir / f"{idx:02d}_{fname}"
        try:
            await _save_upload_to(up, dest)
        finally:
            await up.close()
        archived.append((fname, dest))

    projects.update_project(
        project_id,
        {"status": "running", "stage": "已接收抽签结果，准备更新秩序册赛程", "error": ""},
    )
    threading.Thread(
        target=_run_project_draw_bg,
        args=(project_id, tmp_draw_dir, archived, src_ver["id"], note.strip()),
        daemon=True,
    ).start()
    return projects.get_project(project_id)


def _run_standalone_draw_bg(job_id: str, job_dir: Path, target_copy: Path, target_name: str, draws: list[tuple[str, Path]], note: str) -> None:
    meta_path = job_dir / "job.json"
    try:
        out_docx = job_dir / "updated_book.docx"
        out_pdf = job_dir / "updated_book.pdf"
        res = draw_runner.run_draw_update(
            target_copy,
            draws,
            out_docx,
            out_pdf,
            job_id=f"drawjob_{job_id}",
            note=note,
        )
        storage.write_json(
            meta_path,
            {
                "id": job_id,
                "status": "done",
                "target_name": target_name,
                "created_at": storage.now(),
                "docx_name": out_docx.name,
                "pdf_name": out_pdf.name,
                **res,
            },
        )
    except Exception as exc:
        (job_dir / "draw-error.log").write_text(traceback.format_exc(), encoding="utf-8")
        storage.write_json(
            meta_path,
            {
                "id": job_id,
                "status": "error",
                "target_name": target_name,
                "error": f"{type(exc).__name__}: {exc}",
            },
        )


@router.post("/draw-update", status_code=202)
async def api_standalone_draw_update(
    book_file: UploadFile = File(...),
    draw_files: list[UploadFile] = File(...),
    note: str = Form(""),
):
    """Standalone post-draw update endpoint accepting an uploaded .docx order book + draw result table(s)."""
    book_name = _safe_filename(book_file.filename or "", "秩序册.docx")
    if Path(book_name).suffix.lower() != ".docx":
        raise HTTPException(422, "目标秩序册目前需为 .docx 格式，以便保留原稿排版进行就地替换。")
    if not draw_files:
        raise HTTPException(422, "请至少上传一份抽签结果文件。")

    draw_runner.DRAW_JOBS_DIR.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex[:12]
    jdir = draw_runner.DRAW_JOBS_DIR / job_id
    jdir.mkdir(parents=True, exist_ok=True)

    target_copy = jdir / f"original_{book_name}"
    try:
        await _save_upload_to(book_file, target_copy)
    finally:
        await book_file.close()

    staged_draws: list[tuple[str, Path]] = []
    for idx, up in enumerate(draw_files, 1):
        fname = _safe_filename(up.filename or "", f"draw_{idx}")
        suffix = Path(fname).suffix.lower()
        if suffix not in BOOK_MATERIAL_EXTENSIONS:
            shutil.rmtree(jdir, ignore_errors=True)
            raise HTTPException(422, f"{fname} 的格式不受支持。")
        dest = jdir / f"draw_{idx:02d}_{fname}"
        try:
            await _save_upload_to(up, dest)
        finally:
            await up.close()
        staged_draws.append((fname, dest))

    initial = {
        "id": job_id,
        "status": "running",
        "target_name": book_name,
        "created_at": storage.now(),
    }
    storage.write_json(jdir / "job.json", initial)
    threading.Thread(
        target=_run_standalone_draw_bg,
        args=(job_id, jdir, target_copy, book_name, staged_draws, note.strip()),
        daemon=True,
    ).start()
    return initial


@router.get("/draw-jobs/{job_id}")
def api_get_draw_job(job_id: str):
    meta = storage.read_json(draw_runner.DRAW_JOBS_DIR / job_id / "job.json")
    if not isinstance(meta, dict):
        raise HTTPException(404, "抽签更新任务不存在。")
    return meta


@router.get("/draw-jobs/{job_id}/download/docx")
def api_draw_job_docx(job_id: str):
    jdir = draw_runner.DRAW_JOBS_DIR / job_id
    meta = storage.read_json(jdir / "job.json")
    if not isinstance(meta, dict):
        raise HTTPException(404, "任务不存在。")
    path = jdir / str(meta.get("docx_name") or "updated_book.docx")
    if not path.is_file():
        raise HTTPException(404, "更新后的 Word 文件尚未生成。")
    stem = Path(meta.get("target_name") or "秩序册").stem
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{stem}-抽签更新版.docx",
    )


@router.get("/draw-jobs/{job_id}/download/pdf")
def api_draw_job_pdf(job_id: str):
    jdir = draw_runner.DRAW_JOBS_DIR / job_id
    meta = storage.read_json(jdir / "job.json")
    if not isinstance(meta, dict):
        raise HTTPException(404, "任务不存在。")
    path = jdir / str(meta.get("pdf_name") or "updated_book.pdf")
    if not path.is_file():
        raise HTTPException(404, "更新后的 PDF 文件尚未生成。")
    stem = Path(meta.get("target_name") or "秩序册").stem
    return FileResponse(path, media_type="application/pdf", filename=f"{stem}-抽签更新版.pdf")


@router.get("/draw-jobs/{job_id}/preview/pdf")
def api_draw_job_preview_pdf(job_id: str):
    jdir = draw_runner.DRAW_JOBS_DIR / job_id
    meta = storage.read_json(jdir / "job.json")
    if not isinstance(meta, dict):
        raise HTTPException(404, "任务不存在。")
    path = jdir / str(meta.get("pdf_name") or "updated_book.pdf")
    if not path.is_file():
        raise HTTPException(404, "更新后的 PDF 文件尚未生成。")
    return FileResponse(path, media_type="application/pdf", content_disposition_type="inline")
