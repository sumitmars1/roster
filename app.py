"""FastAPI shell for the Harness-centered roster app."""
from __future__ import annotations

import shutil
import threading
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from starlette.concurrency import run_in_threadpool

import documents
import harness_runner
import storage
from book_api import router as book_router
from config import (
    API_KEY, MAX_FILE_BYTES, MAX_FILES, MAX_REVIEW_ANSWER, MAX_TOTAL_BYTES, STATIC, TASKS,
)
from presets import PRESETS

app = FastAPI(title="体育赛事秩序册整册生成平台", docs_url="/docs")
app.mount("/static", StaticFiles(directory=STATIC), name="static")
app.include_router(book_router)
TASKS.mkdir(parents=True, exist_ok=True)


class ReviewAnswer(BaseModel):
    answer: str


def _task_or_404(task_id: str) -> dict:
    try:
        return storage.read_task(task_id)
    except KeyError:
        raise HTTPException(404, "任务不存在。") from None


def _safe_name(value: str, fallback: str) -> str:
    value = Path(value or "").name.strip()
    return value[:180] or fallback


def _spa_index() -> FileResponse:
    dist_index = STATIC / "dist" / "index.html"
    if dist_index.is_file():
        return FileResponse(dist_index)
    return FileResponse(STATIC / "index.html")


@app.get("/", include_in_schema=False)
def home():
    return _spa_index()


@app.get("/api/options")
def options():
    return {
        "upload": {
            "extensions": sorted(documents.EXTENSIONS),
            "max_files": MAX_FILES,
            "max_file_bytes": MAX_FILE_BYTES,
            "max_total_bytes": MAX_TOTAL_BYTES,
        },
        "presets": [
            {"id": key, "name": value[0], "description": value[1]}
            for key, value in PRESETS.items()
        ],
        "key_configured": bool(API_KEY),
    }


@app.get("/api/history")
def history():
    return storage.history()


@app.post("/api/tasks", status_code=201)
async def create_task(
    title: str = Form(...),
    group: str = Form(""),
    preset: str = Form(...),
    files: list[UploadFile] = File(...),
):
    title, group = title.strip(), group.strip()
    if not title or len(title) > 100:
        raise HTTPException(422, "赛事名称必填，且不能超过 100 字。")
    if len(group) > 70:
        raise HTTPException(422, "组别不能超过 70 字。")
    if preset not in PRESETS:
        raise HTTPException(422, "请选择有效的名单版式。")
    if not 1 <= len(files) <= MAX_FILES:
        raise HTTPException(422, f"一次需要 1–{MAX_FILES} 份资料。")

    task_id = storage.new_task_id()
    directory = TASKS / task_id
    directory.mkdir(parents=True)
    saved = []
    total = 0
    try:
        for index, upload in enumerate(files, 1):
            original = _safe_name(upload.filename or "", f"资料{index}")
            suffix = Path(original).suffix.lower()
            if suffix not in documents.EXTENSIONS:
                raise HTTPException(422, f"{original} 的格式不受支持。")
            stored_name = f"f{index:02d}{suffix}"
            path = directory / stored_name
            size = 0
            with path.open("wb") as handle:
                while chunk := await upload.read(1024 * 1024):
                    size += len(chunk)
                    total += len(chunk)
                    if size > MAX_FILE_BYTES:
                        raise HTTPException(422, f"{original} 超过单份 300 MB 上限。")
                    if total > MAX_TOTAL_BYTES:
                        raise HTTPException(422, "整批资料超过 3 GB 上限。")
                    handle.write(chunk)
            if not size:
                raise HTTPException(422, f"{original} 是空文件。")
            saved.append({
                "id": f"F{index}",
                "name": original,
                "stored_name": stored_name,
                "suffix": suffix,
                "size": size,
                "status": "waiting",
            })
        for file in saved:
            path = directory / file["stored_name"]
            try:
                units = await run_in_threadpool(documents.inventory, path, bulk=True)
            except (OSError, ValueError) as exc:
                raise HTTPException(422, f'{file["name"]} 读取失败：{exc}') from exc
            file["units"] = units
    except HTTPException:
        _remove(directory)
        raise
    except Exception:
        _remove(directory)
        raise
    finally:
        for upload in files:
            await upload.close()

    state = {
        "id": task_id,
        "title": title,
        "group": group,
        "preset": preset,
        "preset_name": PRESETS[preset][0],
        "status": "queued",
        "files": saved,
        "session_id": f"task-{task_id}",
        "usage": {"input_tokens": 0, "output_tokens": 0},
        "result": None,
        "review": None,
        "error": None,
        "answers": [],
        "created_at": storage.now(),
        "updated_at": storage.now(),
        "started_at": None,
        "finished_at": None,
    }
    storage.write_json(directory / "state.json", state)
    threading.Thread(target=harness_runner.run_task, args=(task_id,), daemon=True).start()
    return storage.public(task_id)


def _remove(directory: Path) -> None:
    import shutil

    shutil.rmtree(directory, ignore_errors=True)


@app.post("/api/tasks/{task_id}/files/{file_id}/retry", status_code=202)
def retry_file(task_id: str, file_id: str):
    _task_or_404(task_id)

    def claim(state):
        file = next((item for item in state.get("files", []) if item.get("id") == file_id), None)
        if file is None:
            raise HTTPException(404, "资料不存在。")
        if state.get("status") != "error" or file.get("status") != "error":
            raise HTTPException(409, "只能在任务结束后重试失败的资料。")
        state.update(status="running", error=None, review=None, finished_at=None)

    storage.mutate_task(task_id, claim)
    threading.Thread(target=harness_runner.retry_file, args=(task_id, file_id), daemon=True).start()
    return storage.public(task_id)


@app.post("/api/tasks/{task_id}/restart", status_code=201)
def restart(task_id: str):
    state = _task_or_404(task_id)
    if state.get("status") != "error":
        raise HTTPException(409, "只有失败的任务可以重新整理。")
    old_dir = TASKS / task_id
    new_id = storage.new_task_id()
    directory = TASKS / new_id
    directory.mkdir(parents=True)
    files = []
    try:
        for file in state.get("files", []):
            source = old_dir / file["stored_name"]
            if not source.is_file():
                raise HTTPException(409, f'{file["name"]} 的原文件已不存在，请重新上传。')
            shutil.copy2(source, directory / file["stored_name"])
            copied = dict(file)
            copied["status"] = "waiting"
            copied.pop("error", None)
            copied.pop("session_id", None)
            files.append(copied)
    except Exception:
        _remove(directory)
        raise
    new_state = {
        "id": new_id, "title": state.get("title"), "group": state.get("group", ""),
        "preset": state["preset"], "preset_name": state.get("preset_name"),
        "status": "queued", "files": files, "session_id": f"task-{new_id}",
        "usage": {"input_tokens": 0, "output_tokens": 0}, "result": None, "review": None,
        "error": None, "answers": [], "active_file": None, "created_at": storage.now(),
        "updated_at": storage.now(), "started_at": None, "finished_at": None,
    }
    storage.write_json(directory / "state.json", new_state)
    threading.Thread(target=harness_runner.run_task, args=(new_id,), daemon=True).start()
    return storage.public(new_id)


@app.get("/api/tasks/{task_id}")
def task(task_id: str):
    _task_or_404(task_id)
    return storage.public(task_id)


@app.post("/api/tasks/{task_id}/review", status_code=202)
def review(task_id: str, body: ReviewAnswer):
    _task_or_404(task_id)
    answer = body.answer.strip()
    if not answer:
        raise HTTPException(422, "请填写核对答复。")
    if len(answer) > MAX_REVIEW_ANSWER:
        raise HTTPException(422, f"核对答复不能超过 {MAX_REVIEW_ANSWER} 字。")
    def claim(state):
        if state.get("status") != "review":
            raise HTTPException(409, "这个任务当前不需要核对。")
        answers = list(state.get("answers") or [])
        answers.append({
            "question": (state.get("review") or {}).get("question", ""),
            "answer": answer,
            "at": storage.now(),
        })
        state.update(status="running", review=None, error=None, answers=answers, finished_at=None)

    storage.mutate_task(task_id, claim)
    threading.Thread(target=harness_runner.run_task, args=(task_id,),
                     kwargs={"resume": True, "answer": answer}, daemon=True).start()
    return storage.public(task_id)


@app.get("/api/tasks/{task_id}/download/word")
def download_word(task_id: str):
    state = _task_or_404(task_id)
    result = state.get("result") or {}
    path = TASKS / task_id / str(result.get("word") or "")
    if not path.is_file():
        raise HTTPException(404, "Word 文件不存在。")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{state.get('title') or '名单'}-{state.get('group') or ''}名单.docx".replace("--", "-"),
    )


@app.get("/api/tasks/{task_id}/download/placement")
def download_placement(task_id: str):
    state = _task_or_404(task_id)
    path = TASKS / task_id / "placement.json"
    if not path.is_file():
        raise HTTPException(404, "落位报告不存在。")
    return FileResponse(path, media_type="application/json",
                        filename=f"{state.get('title') or '名单'}落位.json")


@app.get("/api/tasks/{task_id}/source/{file_id}")
def source(task_id: str, file_id: str):
    _task_or_404(task_id)
    for file in storage.read_task(task_id).get("files", []):
        if file.get("id") == file_id:
            path = TASKS / task_id / file["stored_name"]
            if path.is_file():
                return FileResponse(path, filename=file["name"])
    raise HTTPException(404, "资料不存在。")


@app.get("/{spa_path:path}", include_in_schema=False)
def spa_fallback(spa_path: str):
    if spa_path.startswith(("api/", "static/")) or spa_path in {"api", "static", "docs", "openapi.json", "redoc"}:
        raise HTTPException(404, "Not Found")
    return _spa_index()
