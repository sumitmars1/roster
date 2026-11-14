"""MCP tools exposed to the DeepSeek Harness agent.

This process owns document observations and the final submission. Validation
errors are returned as tool errors: the Harness loop can read them, change its
approach, and continue in the same session.

Large rosters are submitted in chunks.  A single giant tool-call JSON is the
most common avoidable failure for an 80-person spreadsheet; ``append_rows`` and
``append_roster`` let the model keep each observation small and use
``finish_roster`` once at the end.
"""
from __future__ import annotations

import argparse
import asyncio
import base64
import json
import logging
from pathlib import Path
from typing import Any

from mcp import types
from mcp.server.lowlevel import Server
from mcp.server.stdio import stdio_server

import document_bank
import storage
from renderer import RenderError, render
from roster_validation import RosterError, source_matches, validate, validate_content, walk_records

MAX_NOTE = 2000
MAX_QUESTION = 1200
MAX_CHUNK_BYTES = 30_000
MAX_APPEND_ROWS = 40
ROW_PRESETS = {"individual-table", "paired-table", "unit-contact-table"}
SCOPE_FILES: tuple[str, ...] = ()


def text_result(text: str, *, error: bool = False) -> types.CallToolResult:
    return types.CallToolResult(isError=error, content=[types.TextContent(type="text", text=text)])


def setup(task_id: str) -> None:
    directory = storage.task_dir(task_id)
    logging.basicConfig(
        filename=directory / "mcp.log", level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8",
    )


def list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="list_documents",
            description=(
                "查看本次任务的全部资料目录。返回每个文件的编号、文件名、页/工作表/正文单元，"
                "以及每个单元是否已经读过。先调用本工具，不要凭文件名猜内容。"
            ),
            inputSchema={"type": "object", "properties": {}, "additionalProperties": False},
        ),
        types.Tool(
            name="read_document",
            description=(
                "读取任意资料、任意页或行范围。PDF 页可用 text（文字）、image（原页图）、"
                "layout（文字坐标）三种视图；text 从 start 行开始读 limit 行，有 next_start 就可以接着读。"
                "PDF 返回 needs_image 时用 image 查看该页，Excel/Word 使用 text。"
                "可以反复回读前面的页或已读范围，用来补队名、组别、跨页号码等字段。"
                "工具返回的 rows[].source 必须原样写进最终名单。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "list_documents 返回的文件编号，例如 F1。"},
                    "unit": {"type": "integer", "minimum": 1, "description": "该文件的页、工作表或正文单元编号。"},
                    "start": {"type": "integer", "minimum": 1, "description": "从 1 开始的文字行或表格行位置。"},
                    "limit": {"type": "integer", "minimum": 1, "maximum": 200, "description": "读取行数，默认 60。"},
                    "view": {"type": "string", "enum": ["text", "image", "layout"],
                             "description": "text 文字；image 原页图；layout PDF 文字坐标。"},
                },
                "required": ["file_id", "unit"],
                "additionalProperties": False,
            },
        ),
        types.Tool(
            name="append_rows",
            description=(
                "逐人/配对/单位联系人母版的分批提交入口。一次提交一个分段标题下的若干行，"
                "建议每次 20–30 行；同名分段会自动合并。每行必须是母版要求的对象，且带 source。"
                "全部行提交完后调用 finish_roster。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "title": {"type": "string", "description": "分段标题，例如组别名或赛事名。"},
                    "group": {"type": "string", "description": "未单独指定组别的行使用此组名。"},
                    "rows": {"type": "array", "items": {"type": "object"},
                             "description": "该分段的行对象数组，字段按母版示例。"},
                },
                "required": ["rows"],
                "additionalProperties": False,
            },
        ),
        types.Tool(
            name="append_roster",
            description=(
                "横排/团体/双栏/分组等结构化母版的分批提交入口。一次提交一个 section（含 data）"
                "或一个 group（含 teams），可以多次调用。双栏编号的运动员放在 "
                "entry.teams[].players，续段沿用相同队名。所有分段提交完后调用 finish_roster。"
                "不要一次提交整份很大的嵌套 JSON。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry": {"type": "object", "description": "一个 section（含 data）或一个 group（含 teams）。"},
                },
                "required": ["entry"],
                "additionalProperties": False,
            },
        ),
        types.Tool(
            name="finish_roster",
            description=(
                "结束分批提交并校验读取范围与 source。多文件任务保存本组片段，单会话任务生成 Word。"
                "校验失败会返回具体原因；补读或修正后可再次 append 再 finish，不要放弃。"
            ),
            inputSchema={
                "type": "object",
                "properties": {"note": {"type": "string", "description": "给用户的简短说明。"}},
                "additionalProperties": False,
            },
        ),
        types.Tool(
            name="submit_roster",
            description=(
                "一次性提交完整名单（只适合较小名单），或提交 status=review 的人工核对问题。"
                "较大的名单请改用 append_rows/append_roster + finish_roster。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "status": {"type": "string", "enum": ["ok", "review"]},
                    "roster": {"type": "object", "description": "按任务指定母版形状填写的完整名单对象。"},
                    "question": {"type": "string", "description": "status=review 时给用户看的具体问题。"},
                    "options": {"type": "array", "items": {"type": "string"},
                                "description": "可选的 2–6 个简短回复选项。"},
                    "sources": {"type": "array", "items": {"type": "string"},
                                "description": "问题涉及的已读来源编号。"},
                    "note": {"type": "string", "description": "给用户的简短说明。"},
                },
                "required": ["status"],
                "additionalProperties": False,
            },
        ),
    ]


def _task(task_id: str) -> dict:
    return storage.read_task(task_id)


def _scope_id() -> str:
    return SCOPE_FILES[0] if SCOPE_FILES else ""


def _parts_path(task_id: str) -> Path:
    directory = storage.task_dir(task_id)
    return directory / "parts" / f"{_scope_id()}.json" if SCOPE_FILES else directory / "parts.json"


def _fragment_path(task_id: str) -> Path:
    return storage.task_dir(task_id) / "fragments" / f"{_scope_id()}.json"


def _source_anchors(task_id: str) -> dict:
    anchors = document_bank.anchors(task_id)
    if not SCOPE_FILES:
        return anchors
    return {source: meta for source, meta in anchors.items() if meta.get("file") in SCOPE_FILES}


def _parts(task_id: str) -> list[dict]:
    value = storage.read_json(_parts_path(task_id), [])
    return value if isinstance(value, list) else []


def _save_parts(task_id: str, parts: list[dict]) -> None:
    storage.write_json(_parts_path(task_id), parts)


def _ensure_not_done(task_id: str) -> None:
    if SCOPE_FILES and _fragment_path(task_id).is_file():
        raise RosterError(f"{_scope_id()} 的文件组片段已经保存，不需要继续提交。")
    if not SCOPE_FILES and (storage.task_dir(task_id) / "result.json").is_file():
        raise RosterError("本任务已经生成 Word，不需要继续提交。")


def _check_sources(rows: list[dict], anchors: dict) -> None:
    for index, row in enumerate(rows, 1):
        if not isinstance(row, dict):
            raise RosterError(f"第 {index} 行不是对象。")
        source = row.get("source")
        if not isinstance(source, str) or not source.strip():
            name = row.get("name") or row.get("team") or row.get("contact") or ""
            raise RosterError(
                f"第 {index} 行（{str(name)[:40]}）缺少 source。"
                "请原样使用读取工具返回的 source。"
            )
        if not source_matches(source, anchors):
            examples = "、".join(list(anchors)[:8]) or "none"
            raise RosterError(
                f"第 {index} 行的 source 不是读取工具返回的编号：{source[:100]}。"
                f"请先读取该页/行；可用示例：{examples}。"
            )


def _append_rows(task_id: str, args: dict) -> str:
    task = _task(task_id)
    preset = task.get("preset")
    if preset not in ROW_PRESETS:
        raise RosterError("append_rows 只用于逐人、配对或单位联系人母版；当前母版请用 append_roster。")
    rows = args.get("rows")
    if not isinstance(rows, list) or not rows:
        raise RosterError("rows 必须是非空数组。")
    if len(rows) > MAX_APPEND_ROWS:
        raise RosterError(f"每次最多 append {MAX_APPEND_ROWS} 行；请分多次提交。")
    if len(json.dumps(rows, ensure_ascii=False).encode()) > MAX_CHUNK_BYTES:
        raise RosterError("本次 rows 太大；请减少行数，分多次提交。")
    validate_content(rows, preset)
    _check_sources(rows, _source_anchors(task_id))
    default = str(args.get("group") or args.get("title") or task.get("group") or task.get("title") or "名单").strip()
    row_key = "pairs" if preset == "paired-table" else "rows"
    buckets: dict[str, list[dict]] = {}
    for row in rows:
        row = dict(row)
        group = str(row.pop("group", "") or row.pop("section", "") or default).strip() or default
        buckets.setdefault(group, []).append(row)
    with storage.locked(storage.task_dir(task_id) / "parts.lock"):
        parts = _parts(task_id)
        for title, bucket in buckets.items():
            merged = False
            for entry in parts:
                data = entry.get("data") if isinstance(entry, dict) else None
                if not isinstance(data, dict) or data.get("title") != title:
                    continue
                if row_key == "pairs" and row_key not in data and "rows" in data:
                    data[row_key] = data.pop("rows")
                if isinstance(data.get(row_key), list):
                    data[row_key].extend(bucket)
                    merged = True
                    break
            if not merged:
                parts.append({"data": {"title": title, row_key: bucket}})
        _save_parts(task_id, parts)
    count = sum(len(entry.get("data", {}).get(row_key, entry.get("data", {}).get("rows", [])))
                for entry in parts if isinstance(entry, dict))
    labels = "、".join(f"{title} {len(bucket)} 行" for title, bucket in buckets.items())
    storage.append_trace(task_id, {"kind": "tool", "name": "append_rows",
                                   "detail": f"{labels} · 累计 {count} 行"})
    return f"已接收 {len(rows)} 行（{labels}），累计 {count} 行。继续读取或提交；全部完成后调用 finish_roster。"


def _append_roster(task_id: str, args: dict) -> str:
    task = _task(task_id)
    preset = task.get("preset")
    entry = args.get("entry")
    if not isinstance(entry, dict) or not entry:
        raise RosterError("entry 必须是一个非空对象。")
    if len(json.dumps(entry, ensure_ascii=False).encode()) > MAX_CHUNK_BYTES:
        raise RosterError("本次 entry 太大；请拆成一个组或一个队再提交。")
    if preset == "two-column-numbered":
        if not isinstance(entry.get("teams"), list) or not entry["teams"]:
            raise RosterError(
                "双栏编号需要 entry: {teams: [{name: 队名, players: [...] }]}。"
                "请把运动员放在 entry.teams[].players，续段沿用同一队名。"
            )
    elif not isinstance(entry.get("data"), dict):
        raise RosterError("当前母版每次提交一个 section 对象，必须包含 data。")
    validate_content(entry, preset)
    _check_sources(list(walk_records(entry)), _source_anchors(task_id))
    with storage.locked(storage.task_dir(task_id) / "parts.lock"):
        parts = _parts(task_id)
        parts.append(entry)
        _save_parts(task_id, parts)
    storage.append_trace(task_id, {"kind": "tool", "name": "append_roster",
                                   "detail": f"已添加第 {len(parts)} 个分段"})
    return f"已添加第 {len(parts)} 个分段。全部提交完后调用 finish_roster。"


def _finish(task_id: str, roster: dict, note: str = "") -> str:
    directory = storage.task_dir(task_id)
    task = _task(task_id)
    progress = document_bank.progress(task_id)
    missing = [
        f'{file["id"]} {unit.get("label", unit.get("index"))}'
        for file in task.get("files", [])
        if not SCOPE_FILES or file["id"] in SCOPE_FILES
        for unit in file.get("units", [])
        if unit["index"] not in progress.get(file["id"], [])
    ]
    if missing:
        raise RosterError(f"还有内容未读完：{'、'.join(missing[:8])}。请读完这些页或行范围，再提交名单。")
    checked = validate(roster, task["preset"], _source_anchors(task_id))
    if SCOPE_FILES:
        storage.write_json(_fragment_path(task_id), {
            "file": _scope_id(), "files": list(SCOPE_FILES), "records": checked["records"],
            "note": note, "created_at": storage.now(), "roster": roster,
        })
        (directory / "review.json").unlink(missing_ok=True)
        storage.append_trace(task_id, {"kind": "done", "name": "finish_fragment",
                                       "detail": f"{_scope_id()} 文件组已保存，共 {checked['records']} 条记录"})
        return json.dumps({
            "ok": True, "files": list(SCOPE_FILES), "records": checked["records"],
            "message": f"{_scope_id()} 文件组片段已保存。不要读取其他文件组，结束本轮。",
        }, ensure_ascii=False)
    if (directory / "result.json").is_file():
        return "本任务已经生成 Word，无需重复提交。"
    info = render(task_id, directory, task["preset"], roster,
                  str(task.get("title") or ""), str(task.get("group") or ""))
    storage.write_json(directory / "result.json", {
        "kind": "done", "records": checked["records"], "note": note,
        "created_at": storage.now(), **info,
    })
    (directory / "review.json").unlink(missing_ok=True)
    storage.append_trace(task_id, {"kind": "done", "name": "finish_roster",
                                   "detail": f"已生成 Word，共 {checked['records']} 条记录"})
    return json.dumps({
        "ok": True, "records": checked["records"],
        "message": "名单已通过校验，Word 已生成。可以结束本轮。",
    }, ensure_ascii=False)


def _finish_parts(task_id: str, note: str = "") -> str:
    parts = _parts(task_id)
    if not parts:
        raise RosterError("还没有提交任何分段；请先 append_rows 或 append_roster。")
    task = _task(task_id)
    roster = {"groups": parts} if task.get("preset") == "two-column-numbered" else {"sections": parts}
    with storage.locked(storage.task_dir(task_id) / "submit.lock"):
        return _finish(task_id, roster, note)


def _submit(task_id: str, args: dict) -> str:
    directory = storage.task_dir(task_id)
    status = args.get("status")
    if status not in {"ok", "review"}:
        raise RosterError("status 只能是 ok 或 review。")
    if status == "ok":
        if _parts(task_id):
            raise RosterError("已经有分批提交的分段；请调用 finish_roster，不要再用 submit_roster(status=ok) 覆盖。")
        roster = args.get("roster")
        if not isinstance(roster, dict):
            raise RosterError("status=ok 时必须在 roster 里提供完整名单对象。")
        with storage.locked(directory / "submit.lock"):
            return _finish(task_id, roster, str(args.get("note") or "").strip()[:MAX_NOTE])

    note = str(args.get("note") or "").strip()[:MAX_NOTE]
    question = str(args.get("question") or "").strip()
    if not question:
        raise RosterError("status=review 时必须在 question 里写清楚需要用户核对什么。")
    if len(question) > MAX_QUESTION:
        raise RosterError(f"核对问题不能超过 {MAX_QUESTION} 字。")
    options = args.get("options") or []
    if not isinstance(options, list) or any(not isinstance(item, str) for item in options):
        raise RosterError("options 必须是字符串数组。")
    options = [item.strip() for item in options if item.strip()]
    if options and not 2 <= len(options) <= 6:
        raise RosterError("options 要么省略，要么提供 2–6 个简短选项。")
    if len(set(options)) != len(options):
        raise RosterError("options 不能重复。")
    if any(len(item) > 300 for item in options):
        raise RosterError("单个选项不能超过 300 字。")
    sources = args.get("sources") or []
    if not isinstance(sources, list) or any(not isinstance(item, str) for item in sources):
        raise RosterError("sources 必须是字符串数组。")
    known = _source_anchors(task_id)
    invalid = [item for item in sources if item not in known]
    if invalid:
        raise RosterError(f"这些 sources 不是读取工具返回的编号：{invalid[:5]}。请原样引用已读来源。")
    storage.write_json(directory / "review.json", {
        "question": question, "options": options, "sources": list(dict.fromkeys(sources)),
        "note": note, "created_at": storage.now(),
    })
    storage.append_trace(task_id, {"kind": "review", "name": "submit_roster", "detail": question})
    return "已记录待核对问题，本轮可以结束；用户答复后会回到同一会话。"


async def serve(task_id: str, files: tuple[str, ...] = ()) -> None:
    global SCOPE_FILES
    SCOPE_FILES = tuple(files)
    setup(task_id)

    async def on_list_tools(_context, _params):
        return types.ListToolsResult(tools=list_tools())

    async def on_call_tool(_context, params):
        name = params.name
        args = params.arguments or {}
        try:
            _ensure_not_done(task_id)
            if name == "list_documents":
                value = document_bank.list_documents(task_id)
                if SCOPE_FILES:
                    value["files"] = [item for item in value["files"] if item.get("id") in SCOPE_FILES]
                storage.append_trace(task_id, {"kind": "tool", "name": name,
                                               "detail": f"{len(value['files'])} 份资料"})
                return text_result(json.dumps(value, ensure_ascii=False))
            if name == "read_document":
                if SCOPE_FILES and args.get("file_id") not in SCOPE_FILES:
                    raise RosterError(f"本轮只处理 {', '.join(SCOPE_FILES)}；不要读取其他资料。")
                payload, image = document_bank.read_document(task_id, args)
                detail = f"{payload.get('file')} · {payload.get('label') or payload.get('unit')} · {payload.get('view')}"
                storage.append_trace(task_id, {"kind": "tool", "name": name, "detail": detail})
                content = [types.TextContent(type="text", text=json.dumps(payload, ensure_ascii=False))]
                if image:
                    content.append(types.ImageContent(
                        type="image", data=base64.b64encode(image).decode("ascii"), mimeType="image/jpeg",
                    ))
                return types.CallToolResult(content=content)
            if name == "append_rows":
                return text_result(_append_rows(task_id, args))
            if name == "append_roster":
                return text_result(_append_roster(task_id, args))
            if name == "finish_roster":
                return text_result(_finish_parts(task_id, str(args.get("note") or "").strip()[:MAX_NOTE]))
            if name == "submit_roster":
                return text_result(_submit(task_id, args))
            return text_result(f"未知工具：{name}", error=True)
        except (document_bank.DocumentError, RosterError, RenderError, OSError, ValueError) as exc:
            message = str(exc).strip() or "工具执行失败。"
            storage.append_trace(task_id, {"kind": "tool", "name": name, "error": message[:600]})
            return text_result(message, error=True)
        except Exception:
            logging.exception("tool failed: %s", name)
            storage.append_trace(task_id, {"kind": "tool", "name": name, "error": "工具内部错误，请换一种读取方式。"})
            return text_result("工具内部错误。请换一种读取方式，或缩小读取范围。", error=True)

    server = Server("roster", on_list_tools=on_list_tools, on_call_tool=on_call_tool)
    async with stdio_server() as (reader, writer):
        await server.run(reader, writer, server.create_initialization_options())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--task", required=True)
    parser.add_argument("--files", help="comma-separated scoped file IDs")
    args = parser.parse_args()
    files = tuple(item for item in (args.files or "").split(",") if item)
    try:
        asyncio.run(serve(args.task, files))
    except KeyboardInterrupt:
        pass
    except Exception:
        logging.exception("MCP server stopped")
        raise


if __name__ == "__main__":
    main()
