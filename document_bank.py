"""Document access for one roster task.

The tool layer is deliberately a library of observations, not a batch state
machine: the model may read any page, cell range, or image in any order.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pymupdf

import documents
from config import MAX_FILE_BYTES
from limits import MAX_TOOL_BYTES
from storage import mutate_json, read_json, task_dir

DEFAULT_ROWS = 60
MAX_ROWS = 200


class DocumentError(ValueError):
    pass


def state(task_id: str) -> dict:
    value = read_json(task_dir(task_id) / "state.json")
    if not isinstance(value, dict):
        raise DocumentError("任务状态不存在。")
    return value


def files(task_id: str) -> list[dict]:
    return state(task_id).get("files", [])


def by_id(task_id: str, file_id: str) -> dict:
    for file in files(task_id):
        if file.get("id") == file_id:
            return file
    raise DocumentError(f"没有编号为 {file_id} 的资料。请先调用 list_documents 查看可用编号。")


def unit_of(file: dict, index: int) -> dict:
    for unit in file.get("units", []):
        if unit.get("index") == index:
            return unit
    raise DocumentError(f"{file['id']} 没有第 {index} 个读取单元。请查看目录中的单元编号。")


def _anchors(task_id: str) -> dict:
    value = read_json(task_dir(task_id) / "anchors.json", {})
    return value if isinstance(value, dict) else {}


def anchors(task_id: str) -> dict:
    return _anchors(task_id)


def _add_anchors(task_id: str, values: dict[str, dict]) -> None:
    if not values:
        return
    path = task_dir(task_id) / "anchors.json"
    mutate_json(path, lambda current: {**(current or {}), **values}, {})


def _mark_read(task_id: str, file_id: str, unit: int, start: int, stop: int, total: int) -> None:
    path = task_dir(task_id) / "progress.json"

    def change(current):
        current = current if isinstance(current, dict) else {}
        entries = current.get(file_id)
        if not isinstance(entries, dict):
            entries = current[file_id] = {}
        key = str(unit)
        if entries.get(key) is True:
            return current
        ranges = sorted([*entries.get(key, []), [start, stop]])
        merged = []
        for first, last in ranges:
            if merged and first <= merged[-1][1] + 1:
                merged[-1][1] = max(last, merged[-1][1])
            else:
                merged.append([first, last])
        entries[key] = True if merged[0][0] == 1 and merged[0][1] >= total else merged
        return current

    mutate_json(path, change, {})


def progress(task_id: str) -> dict:
    value = read_json(task_dir(task_id) / "progress.json", {})
    if not isinstance(value, dict):
        return {}
    return {
        file_id: [int(index) for index, complete in units.items() if complete is True]
        for file_id, units in value.items() if isinstance(units, dict)
    }


def reset_file_reads(task_id: str, file_ids: list[str]) -> None:
    directory = task_dir(task_id)
    mutate_json(directory / "progress.json", lambda current: {
        key: value for key, value in current.items() if key not in file_ids
    }, {})
    mutate_json(directory / "anchors.json", lambda current: {
        key: value for key, value in current.items() if value.get("file") not in file_ids
    }, {})


def _excerpt(value: Any) -> str:
    if isinstance(value, list):
        value = " | ".join(str(item) for item in value if str(item).strip())
    text = str(value).replace("\x00", "")
    return text[:1000]


def _source_meta(file: dict, unit: dict) -> dict:
    return {
        "file": file["id"],
        "page": unit["index"] if unit.get("kind") == "page" else None,
        "unit": unit["index"],
    }


def _broad_anchor(file: dict, unit: dict) -> str:
    prefix = f"{file['id']} {file['name']}"
    if unit.get("kind") == "page":
        return f"{prefix} / 第 {unit['index']} 页"
    if unit.get("kind") == "image":
        return f"{prefix} / 图片"
    return f"{prefix} / {unit.get('label') or '正文'}"


def list_documents(task_id: str) -> dict:
    current = progress(task_id)
    result = []
    for file in files(task_id):
        read = set(current.get(file["id"], []))
        result.append({
            "id": file["id"],
            "name": file["name"],
            "suffix": file.get("suffix"),
            "size": file.get("size"),
            "units": [
                {
                    "index": unit["index"],
                    "label": unit["label"],
                    "kind": unit["kind"],
                    "size": unit["size"],
                    "read": unit["index"] in read,
                }
                for unit in file.get("units", [])
            ],
        })
    return {"files": result}


def _layout(task_id: str, file: dict, unit: dict) -> dict:
    if unit.get("kind") != "page":
        raise DocumentError("layout 视图只适用于 PDF 页；Office 表格请用 text 视图。")
    path = task_dir(task_id) / file["stored_name"]
    with pymupdf.open(path) as pdf:
        page = pdf[unit["index"] - 1]
        words = page.get_text("words", sort=True)
        if not words:
            raise DocumentError("本页没有可提取的文字位置。请改用 image 视图查看原页。")
        anchor = f"{file['id']} {file['name']} / 第 {unit['index']} 页文字位置"
        broad = _broad_anchor(file, unit)
        needs_image = documents.page_has_large_image(page)
        payload = {
            "source": anchor,
            "file": file["id"],
            "unit": unit["index"],
            "view": "layout",
            "页宽": round(page.rect.width, 1),
            "页高": round(page.rect.height, 1),
            "文字位置": [[word[4], *[round(value, 1) for value in word[:4]]] for word in words],
            "needs_image": needs_image,
            "说明": "每项为 [文字, x0, y0, x1, y1]，按阅读顺序返回。",
        }
        if needs_image:
            payload["说明"] += " 本页含较大图像，请同时读取 image 视图。"
        if _wire_size(payload) > MAX_TOOL_BYTES:
            raise DocumentError("本页文字位置过多，请改用 text 分段或 image 视图。")
    _add_anchors(task_id, {
        anchor: {**_source_meta(file, unit), "excerpt": "PDF 文字位置"},
        broad: {**_source_meta(file, unit), "excerpt": "PDF 页"},
    })
    if not needs_image:
        _mark_read(task_id, file["id"], unit["index"], 1, 1, 1)
    return payload


def _image(task_id: str, file: dict, unit: dict) -> tuple[dict, bytes]:
    if unit.get("kind") not in {"page", "image"}:
        raise DocumentError("只有 PDF 页和原始图片提供 image 视图；Office 文件请用 text 读取单元格。")
    path = task_dir(task_id) / file["stored_name"]
    prefix = f"{file['id']} {file['name']}"
    anchor, image = documents.page_image(path, unit, prefix)
    broad = _broad_anchor(file, unit)
    _add_anchors(task_id, {
        anchor: {**_source_meta(file, unit), "excerpt": "原页图"},
        broad: {**_source_meta(file, unit), "excerpt": "原页图"},
    })
    _mark_read(task_id, file["id"], unit["index"], 1, 1, 1)
    return {
        "source": anchor,
        "file": file["id"],
        "unit": unit["index"],
        "view": "image",
        "说明": "原页图已随本工具结果返回。可以继续读取其他页，或回读本页文字。",
    }, image


def _text(task_id: str, file: dict, unit: dict, start: int, limit: int) -> dict:
    path = task_dir(task_id) / file["stored_name"]
    prefix = f"{file['id']} {file['name']}"
    is_page = unit["kind"] == "page"
    if start < 1 or (not is_page and start > max(unit["size"], 1)):
        raise DocumentError(f"起始位置 {start} 超出该单元范围。")
    rows, following, needs_image, total = documents.extract(
        path, unit, start, prefix, bulk=True, row_limit=limit,
    )
    if start > 1 and not rows:
        raise DocumentError(f"起始位置 {start} 超出该单元范围。")
    broad = _broad_anchor(file, unit)
    payload = {
        "source": broad,
        "file": file["id"],
        "unit": unit["index"],
        "view": "text",
        "label": unit.get("label"),
        "rows": [],
        "next_start": following,
        "has_more": following is not None,
        "needs_image": needs_image,
        "说明": (
            "rows[].source 可直接作为名单记录的 source；"
            "跨页或跨表字段可继续读取相邻页/行，不要因当前页缺列而猜。"
        ),
    }
    if needs_image:
        payload["说明"] += " 该单元含图像内容，请读取 image 视图。"

    stop = start + len(rows) - 1
    # Reserve room for a continuation cursor even when the entire range fits.
    wire_size = _wire_size({**payload, "next_start": max(stop + 1, 1), "has_more": False})
    for index, (anchor, content) in enumerate(rows, start):
        if not (any(str(cell).strip() for cell in content) if isinstance(content, list) else content.strip()):
            continue
        item = {"source": anchor, "text": content}
        wire_size += _wire_size(item) + (2 if payload["rows"] else 0)
        if wire_size > MAX_TOOL_BYTES:
            if not payload["rows"]:
                raise DocumentError(f"第 {index} 行超过工具输出上限，请缩短原资料中的该行。")
            payload.update(next_start=index, has_more=True)
            stop = index - 1
            break
        payload["rows"].append(item)
    if _wire_size(payload) > MAX_TOOL_BYTES:
        raise DocumentError("读取结果的基本信息超过工具输出上限。")

    neighbours = _neighbours(task_id, file, unit, start, payload["next_start"]) if is_page else {}
    if neighbours and _wire_size({**payload, "adjacent": neighbours}) <= MAX_TOOL_BYTES:
        payload["adjacent"] = neighbours

    meta = _source_meta(file, unit)
    anchor_values = {
        row["source"]: {**meta, "excerpt": _excerpt(row["text"])} for row in payload["rows"]
    }
    if anchor_values:
        anchor_values[broad] = {**meta, "excerpt": unit.get("label")}
    for direction, items in payload.get("adjacent", {}).items():
        page = unit["index"] + (-1 if direction == "before" else 1)
        adjacent = _source_meta(file, {**unit, "index": page})
        anchor_values.update({
            row["source"]: {**adjacent, "excerpt": _excerpt(row["text"])} for row in items
        })
    _add_anchors(task_id, anchor_values)
    if not needs_image:
        _mark_read(task_id, file["id"], unit["index"], start, max(stop, 1), max(total, 1))
    return payload


def _neighbours(task_id: str, file: dict, unit: dict, start: int, following: int | None) -> dict:
    """Give the model one small look across the page boundary, not a whole batch."""
    path = task_dir(task_id) / file["stored_name"]
    prefix = f"{file['id']} {file['name']}"
    index = unit["index"]
    result = {}
    try:
        with pymupdf.open(path) as pdf:
            for direction, page in (("before", index - 1), ("after", index + 1)):
                if direction == "before" and (start != 1 or page < 1):
                    continue
                if direction == "after" and (following is not None or page > len(pdf)):
                    continue
                lines = [(i, line) for i, line in enumerate(
                    pdf[page - 1].get_text("text", sort=True).splitlines(), 1
                ) if line.strip()]
                sample = lines[-8:] if direction == "before" else lines[:8]
                if sample:
                    result[direction] = [
                        {"source": f"{prefix} / 第 {page} 页 / 第 {i} 行", "text": line}
                        for i, line in sample
                    ]
    except (OSError, ValueError, RuntimeError):
        pass
    return result


def _wire_size(payload: dict) -> int:
    return len(json.dumps(payload, ensure_ascii=False).encode("utf-8"))


def read_document(task_id: str, args: dict) -> tuple[dict, bytes | None]:
    file_id = args.get("file_id")
    if not isinstance(file_id, str):
        raise DocumentError("file_id 必须是 list_documents 返回的编号，例如 F1。")
    file = by_id(task_id, file_id)
    try:
        index = int(args.get("unit"))
    except (TypeError, ValueError):
        raise DocumentError("unit 必须是目录中该资料的单元编号（从 1 开始）。") from None
    unit = unit_of(file, index)
    view = args.get("view") or "text"
    if view == "image":
        return _image(task_id, file, unit)
    if view == "layout":
        return _layout(task_id, file, unit), None
    if view != "text":
        raise DocumentError("view 只能是 text、image 或 layout。")
    try:
        start = int(args.get("start") or 1)
    except (TypeError, ValueError):
        raise DocumentError("start 必须是大于等于 1 的整数。") from None
    try:
        limit = int(args.get("limit") or DEFAULT_ROWS)
    except (TypeError, ValueError):
        raise DocumentError("limit 必须是整数。") from None
    if limit < 1:
        raise DocumentError("limit 必须大于 0。")
    limit = min(limit, MAX_ROWS)
    return _text(task_id, file, unit, start, limit), None


def file_path(task_id: str, file_id: str) -> tuple[Path, dict]:
    file = by_id(task_id, file_id)
    return task_dir(task_id) / file["stored_name"], file


def validate_size(path: Path) -> None:
    if path.stat().st_size > MAX_FILE_BYTES:
        raise DocumentError("文件超过单份大小上限。")
