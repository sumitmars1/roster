"""Small JSON task store shared by the web app and the MCP tool process."""
from __future__ import annotations

import fcntl
import json
import os
import tempfile
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterator

from config import TASKS

TERMINAL = {"done", "error", "stopped"}


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def task_dir(task_id: str) -> Path:
    if (not task_id or len(task_id) > 64
            or any(not (char.isalnum() or char in "-_") for char in task_id)):
        raise KeyError(task_id)
    path = TASKS / task_id
    if not path.is_dir():
        raise KeyError(task_id)
    return path


def new_task_id() -> str:
    from uuid import uuid4

    return str(uuid4())


def read_json(path: Path, default: Any = None) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=path.name + ".", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temporary, path)
    finally:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass


@contextmanager
def locked(path: Path) -> Iterator[None]:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+") as handle:
        fcntl.flock(handle, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(handle, fcntl.LOCK_UN)


def mutate_json(path: Path, change: Callable[[Any], Any], default: Any = None) -> Any:
    """Read-modify-write one JSON document under an advisory lock."""
    with locked(path.with_suffix(path.suffix + ".lock")):
        value = read_json(path, default)
        result = change(value)
        write_json(path, result)
        return result


def read_task(task_id: str) -> dict:
    directory = task_dir(task_id)
    state = read_json(directory / "state.json")
    if not isinstance(state, dict):
        raise KeyError(task_id)
    return state


def update_task(task_id: str, **fields: Any) -> dict:
    return mutate_task(task_id, lambda state: state.update(fields))


def mutate_task(task_id: str, update: Callable[[dict], None]) -> dict:
    path = task_dir(task_id) / "state.json"

    def change(state):
        if not isinstance(state, dict):
            raise KeyError(task_id)
        update(state)
        state["updated_at"] = now()
        return state

    return mutate_json(path, change, {})


def append_trace(task_id: str, event: dict) -> None:
    """Append one UI-facing tool event; the MCP process is the only writer."""
    path = task_dir(task_id) / "trace.jsonl"
    record = {"time": now(), **event}
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")


def read_trace(task_id: str, limit: int = 120) -> list[dict]:
    path = task_dir(task_id) / "trace.jsonl"
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except FileNotFoundError:
        return []
    events = []
    for line in lines[-limit:]:
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(event, dict):
            events.append(event)
    return events


def public(task_id: str) -> dict:
    state = read_task(task_id)
    directory = task_dir(task_id)
    result = state.get("result") or {}
    word = directory / result.get("word", "") if result.get("word") else None
    source_files = state.get("files", [])
    state = {
        key: value for key, value in state.items()
        if key not in {"files"}
    }
    state["files"] = [
        {key: value for key, value in file.items() if key != "stored_name"}
        for file in source_files
    ]
    state["trace"] = read_trace(task_id)
    state["word_available"] = bool(word and word.is_file())
    state["placement_available"] = bool((directory / "placement.json").is_file())
    return state


def history(limit: int = 60) -> list[dict]:
    entries = []
    if not TASKS.is_dir():
        return entries
    for directory in TASKS.iterdir():
        if not directory.is_dir():
            continue
        state = read_json(directory / "state.json")
        if not isinstance(state, dict) or state.get("status") != "done":
            continue
        result = state.get("result") or {}
        word = directory / result.get("word", "") if result.get("word") else None
        entries.append({
            "id": state.get("id"),
            "title": state.get("title"),
            "group": state.get("group"),
            "preset": state.get("preset"),
            "preset_name": state.get("preset_name"),
            "generated_at": state.get("finished_at") or state.get("updated_at"),
            "records": result.get("records"),
            "word_available": bool(word and word.is_file()),
        })
    entries.sort(key=lambda item: item.get("generated_at") or "", reverse=True)
    return entries[:limit]
