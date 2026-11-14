from __future__ import annotations

import json
import shutil
import uuid
from pathlib import Path
from typing import Any

import config
from storage import locked, now, read_json, write_json

PROJECTS_DIR = config.DATA / "projects"


def ensure_projects_dir() -> None:
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)


def project_dir(project_id: str) -> Path:
    return PROJECTS_DIR / project_id


def project_json_path(project_id: str) -> Path:
    return project_dir(project_id) / "project.json"


def project_sources_dir(project_id: str) -> Path:
    d = project_dir(project_id) / "sources"
    d.mkdir(parents=True, exist_ok=True)
    return d


def project_version_dir(project_id: str, version_id: str) -> Path:
    d = project_dir(project_id) / "versions" / version_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def project_trace_path(project_id: str) -> Path:
    return project_dir(project_id) / "trace.jsonl"


def create_project(name: str, template_id: str) -> dict[str, Any]:
    ensure_projects_dir()
    pid = uuid.uuid4().hex[:12]
    pdir = project_dir(pid)
    pdir.mkdir(parents=True, exist_ok=True)
    project_sources_dir(pid)
    stamp = now()
    record: dict[str, Any] = {
        "id": pid,
        "name": (name or "").strip() or "未命名赛事秩序册",
        "template_id": template_id,
        "status": "draft",
        "stage": "待上传资料",
        "error": "",
        "materials": [],
        "versions": [],
        "latest_version_id": None,
        "created_at": stamp,
        "updated_at": stamp,
    }
    write_json(project_json_path(pid), record)
    return record


def get_project(project_id: str) -> dict[str, Any] | None:
    path = project_json_path(project_id)
    if not path.is_file():
        return None
    with locked(path):
        return read_json(path)


def list_projects() -> list[dict[str, Any]]:
    ensure_projects_dir()
    items: list[dict[str, Any]] = []
    for folder in PROJECTS_DIR.iterdir():
        if not folder.is_dir():
            continue
        pjson = folder / "project.json"
        if pjson.is_file():
            items.append(read_json(pjson))
    items.sort(key=lambda p: p.get("updated_at") or p.get("created_at") or "", reverse=True)
    return items


def update_project(project_id: str, updates: dict[str, Any]) -> dict[str, Any]:
    path = project_json_path(project_id)
    with locked(path):
        data = read_json(path)
        data.update(updates)
        data["updated_at"] = now()
        write_json(path, data)
        return data


def append_project_trace(project_id: str, event: dict[str, Any]) -> None:
    path = project_trace_path(project_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    record = {"ts": now(), **event}
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")


def read_project_trace(project_id: str, limit: int = 80) -> list[dict[str, Any]]:
    path = project_trace_path(project_id)
    if not path.is_file():
        return []
    items: list[dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            items.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return items[-limit:]


def _safe_filename(name: str) -> str:
    base = Path(name or "file").name.strip()
    return base.replace("/", "_").replace("\\", "_") or f"file-{uuid.uuid4().hex[:6]}"


def add_material_file(
    project_id: str,
    src_path: Path,
    original_name: str,
    *,
    role_hint: str = "auto",
    replace_material_id: str | None = None,
) -> dict[str, Any]:
    """Copy or replace a source file into project's sources/ directory and record metadata."""
    pjson = project_json_path(project_id)
    sdir = project_sources_dir(project_id)
    clean_name = _safe_filename(original_name)
    mid = replace_material_id or uuid.uuid4().hex[:10]
    stored_name = f"{mid}_{clean_name}"
    dest = sdir / stored_name
    shutil.copy2(src_path, dest)

    with locked(pjson):
        data = read_json(pjson)
        materials = list(data.get("materials") or [])
        if replace_material_id:
            kept: list[dict[str, Any]] = []
            for m in materials:
                if m.get("id") == replace_material_id:
                    old_file = sdir / str(m.get("stored_name") or "")
                    if old_file.is_file() and old_file != dest:
                        old_file.unlink()
                else:
                    kept.append(m)
            materials = kept

        entry = {
            "id": mid,
            "name": clean_name,
            "stored_name": stored_name,
            "size": dest.stat().st_size,
            "role_hint": role_hint or "auto",
            "uploaded_at": now(),
        }
        materials.append(entry)
        data["materials"] = materials
        data["updated_at"] = now()
        write_json(pjson, data)
    return entry


def remove_material(project_id: str, material_id: str) -> bool:
    pjson = project_json_path(project_id)
    sdir = project_sources_dir(project_id)
    removed = False
    with locked(pjson):
        data = read_json(pjson)
        kept: list[dict[str, Any]] = []
        for m in data.get("materials") or []:
            if m.get("id") == material_id:
                f = sdir / str(m.get("stored_name") or "")
                if f.is_file():
                    f.unlink()
                removed = True
            else:
                kept.append(m)
        data["materials"] = kept
        data["updated_at"] = now()
        write_json(pjson, data)
    return removed


def material_file_path(project: dict[str, Any], material: dict[str, Any]) -> Path:
    return project_sources_dir(project["id"]) / str(material["stored_name"])


def allocate_version_id(project_id: str) -> str:
    pjson = project_json_path(project_id)
    with locked(pjson):
        data = read_json(pjson)
        versions = data.get("versions") or []
        idx = len(versions) + 1
        existing = {v.get("id") for v in versions}
        while f"v{idx}" in existing:
            idx += 1
        return f"v{idx}"


def record_version(project_id: str, version_info: dict[str, Any]) -> dict[str, Any]:
    pjson = project_json_path(project_id)
    with locked(pjson):
        data = read_json(pjson)
        versions = list(data.get("versions") or [])
        versions.append(version_info)
        data["versions"] = versions
        data["latest_version_id"] = version_info["id"]
        data["status"] = "ready"
        data["stage"] = "已生成（可在线预览与下载）"
        data["error"] = ""
        data["updated_at"] = now()
        write_json(pjson, data)
        return data


def get_version(project: dict[str, Any], version_id: str | None = None) -> dict[str, Any] | None:
    vid = version_id or project.get("latest_version_id")
    if not vid:
        return None
    for v in reversed(project.get("versions") or []):
        if v.get("id") == vid:
            return v
    return None
