from __future__ import annotations

import argparse
import asyncio
import base64
import json
import sys
import tempfile
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn
from mcp import types
from mcp.server.lowlevel import Server
from mcp.server.stdio import stdio_server

import document_bank
import storage
from book_pdf import convert_to_docx
from draw_update import WIN_LOSS_RE

TASK_ID = ""
MODE = "plan"


def _text_result(text: str, *, error: bool = False) -> types.CallToolResult:
    return types.CallToolResult(
        isError=error,
        content=[types.TextContent(type="text", text=text)],
    )


def _trace(kind: str, name: str, **payload: Any) -> None:
    if TASK_ID:
        storage.append_trace(TASK_ID, {"kind": kind, "name": name, **payload})


def _mark_plan_touched(file_id: str) -> None:
    if not TASK_ID or not file_id:
        return
    p = storage.task_dir(TASK_ID) / "plan_reads.json"
    storage.mutate_json(p, lambda cur: sorted(set([*(cur or []), file_id])), [])


def _record_plan_window(
    file_id: str,
    unit_idx: Any,
    view: str,
    start: int,
    limit: int,
) -> tuple[int, bool]:
    """Track Stage-1 reads per (file_id, unit, view) to detect and break scanning loops."""
    key = f"{file_id}:{unit_idx}:{view}"
    stop = start + max(1, limit) - 1
    p = storage.task_dir(TASK_ID) / "plan_read_windows.json"
    state_holder: dict[str, Any] = {"count": 0, "overlapped": False}

    def _mutate(cur: Any) -> dict[str, Any]:
        data = dict(cur or {})
        entry = dict(data.get(key) or {})
        count = int(entry.get("count") or 0) + 1
        windows = list(entry.get("windows") or [])
        overlapped = any(not (stop < int(w[0]) or start > int(w[1])) for w in windows if isinstance(w, list) and len(w) == 2)
        state_holder["count"] = count
        state_holder["overlapped"] = overlapped
        if count <= 5:
            windows.append([start, stop])
            data[key] = {"count": count, "windows": windows[-20:]}
        return data

    storage.mutate_json(p, _mutate, {})
    return state_holder["count"], state_holder["overlapped"]


def _inspect_docx_outline_impl(file_id: str) -> dict[str, Any]:
    try:
        path, _file_info = document_bank.file_path(TASK_ID, file_id)
    except document_bank.DocumentError as exc:
        return {"error": str(exc)}
    suffix = path.suffix.lower()
    if suffix not in {".docx", ".doc", ".wps"}:
        return {"error": f"{path.name} 不是 Word 文档，请使用 read_document。"}
    with tempfile.TemporaryDirectory(prefix="outline-") as tmp_dir:
        norm = convert_to_docx(path, Path(tmp_dir) / "doc.docx")
        doc = Document(str(norm))
        blocks: list[dict[str, Any]] = []
        p_idx = 0
        t_idx = 0
        for idx, child in enumerate(doc.element.body):
            if child.tag == qn("w:p"):
                txt = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
                has_img = (
                    child.find(".//" + qn("a:blip")) is not None
                    or child.find(".//{urn:schemas-microsoft-com:vml}imagedata") is not None
                )
                if txt or has_img:
                    blocks.append(
                        {
                            "block": idx,
                            "type": "image_paragraph" if (has_img and not txt) else ("mixed_paragraph" if has_img else "paragraph"),
                            "p_idx": p_idx,
                            "text": txt[:140],
                            "has_image": has_img,
                        }
                    )
                p_idx += 1
            elif child.tag == qn("w:tbl"):
                rows = child.findall(qn("w:tr"))
                hdr = []
                if rows:
                    hdr = [
                        "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()[:24]
                        for tc in rows[0].findall(qn("w:tc"))[:10]
                    ]
                blocks.append(
                    {
                        "block": idx,
                        "type": "table",
                        "t_idx": t_idx,
                        "rows": len(rows),
                        "headers": hdr,
                    }
                )
                t_idx += 1
        return {
            "file_id": file_id,
            "file_name": (_file_info or {}).get("name") or path.name,
            "total_blocks": len(blocks),
            "blocks": blocks[:250],
        }


def _validate_book_plan(plan: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if not isinstance(plan, dict):
        return ["plan 必须是 JSON 对象。"]
    state = storage.read_task(TASK_ID)
    valid_file_ids = {f["id"] for f in state.get("files") or []}

    if not str(plan.get("event_title") or "").strip():
        errors.append("缺少 event_title（赛事全称）。")

    cover = plan.get("cover") or {}
    if cover.get("file_id") and cover["file_id"] not in valid_file_ids:
        errors.append(f"cover.file_id 无效: {cover['file_id']}")
    back = plan.get("back_cover") or {}
    if back.get("file_id") and back["file_id"] not in valid_file_ids:
        errors.append(f"back_cover.file_id 无效: {back['file_id']}")
    wm = plan.get("watermark") or {}
    if wm.get("file_id") and wm["file_id"] not in valid_file_ids:
        errors.append(f"watermark.file_id 无效: {wm['file_id']}")

    chapters = plan.get("chapters")
    if not isinstance(chapters, list) or not chapters:
        errors.append("chapters 必须是非空数组。")
    else:
        for idx, ch in enumerate(chapters):
            if not isinstance(ch, dict):
                errors.append(f"chapters[{idx}] 必须是对象。")
                continue
            if not str(ch.get("title") or "").strip():
                errors.append(f"chapters[{idx}] 缺少 title。")
            if not str(ch.get("kind") or "").strip():
                errors.append(f"chapters[{idx}] 缺少 kind。")
            for s_idx, src in enumerate(ch.get("sources") or []):
                fid = src.get("file_id")
                if fid and fid not in valid_file_ids:
                    errors.append(f"chapters[{idx}].sources[{s_idx}].file_id 无效: {fid}")

    roster_scope = plan.get("roster_scope") or {}
    for r_idx, item in enumerate(roster_scope.get("sources") or []):
        fid = item.get("file_id")
        if not fid or fid not in valid_file_ids:
            errors.append(f"roster_scope.sources[{r_idx}].file_id 无效或缺失: {fid}")

    # Check that non-image documents in the batch were actually inspected via read/outline
    tdir = storage.task_dir(TASK_ID)
    progress_data = storage.read_json(tdir / "progress.json", {}) or {}
    touched = set(storage.read_json(tdir / "plan_reads.json", []) or [])
    unread_docs: list[str] = []
    for f in state.get("files") or []:
        ext = Path(f["name"]).suffix.lower()
        if ext in {".png", ".jpg", ".jpeg", ".webp"}:
            continue
        fid = f["id"]
        if not progress_data.get(fid) and fid not in touched:
            unread_docs.append(f"{f['name']}({fid})")
    if unread_docs:
        errors.append(
            "尚有以下上传文档一次未读取，请先调用 read_document 核实内容后再提交规划: "
            + ", ".join(unread_docs)
        )
    return errors


def _validate_draw_patches(args: dict[str, Any]) -> tuple[list[str], dict[str, Any] | None]:
    errors: list[str] = []
    patches = args.get("patches")
    if not isinstance(patches, list):
        return ["patches 必须是数组。"], None

    directory = storage.task_dir(TASK_ID)
    inspect_data = storage.read_json(directory / "book_inspect.json") or {}
    tables = inspect_data.get("tables") or []
    tbl_by_idx = {t["table_index"]: t for t in tables}

    for i, p in enumerate(patches):
        if not isinstance(p, dict):
            errors.append(f"patches[{i}] 必须是对象。")
            continue
        tgt = str(p.get("target") or "cell")
        find_t = str(p.get("find") or "")
        repl_t = str(p.get("replace") or "")
        if not find_t or not repl_t:
            errors.append(f"patches[{i}] 缺少 find 或 replace。")
            continue
        if WIN_LOSS_RE.search(find_t):
            errors.append(f"patches[{i}] 试图替换胜负/名次引用 '{find_t}'，按规则禁止修改。")
            continue
        if tgt == "cell":
            t_idx = p.get("table")
            r_idx = p.get("row")
            c_idx = p.get("col")
            t_meta = tbl_by_idx.get(t_idx)
            if not t_meta:
                errors.append(f"patches[{i}] table={t_idx} 不存在于 target book。")
                continue
            if not (isinstance(r_idx, int) and 0 <= r_idx < t_meta["row_count"]):
                errors.append(f"patches[{i}] row={r_idx} 超出 table[{t_idx}] 行数范围。")
            if not (isinstance(c_idx, int) and 0 <= c_idx < t_meta["col_count"]):
                errors.append(f"patches[{i}] col={c_idx} 超出 table[{t_idx}] 列数范围。")
        elif tgt == "paragraph":
            if not isinstance(p.get("paragraph"), int) or p["paragraph"] < 0:
                errors.append(f"patches[{i}] paragraph 索引无效。")
        else:
            errors.append(f"patches[{i}] target 必须是 'cell' 或 'paragraph'。")

    if errors:
        return errors, None
    return [], {
        "patches": patches,
        "preserved_examples": args.get("preserved_examples") or [],
        "summary": str(args.get("summary") or ""),
    }


def _read_document_tool(description: str) -> types.Tool:
    return types.Tool(
        name="read_document",
        description=description,
        inputSchema={
            "type": "object",
            "properties": {
                "file_id": {"type": "string"},
                "unit": {"type": ["integer", "string"]},
                "start": {"type": "integer", "minimum": 1},
                "limit": {"type": "integer", "minimum": 1, "maximum": 200},
                "view": {"type": "string", "enum": ["text", "image", "layout"]},
                "mode": {"type": "string", "enum": ["text", "image", "layout"]},
            },
            "required": ["file_id"],
            "additionalProperties": False,
        },
    )


def _plan_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="list_documents",
            description="列出本次项目上传的全部原始资料文件及其 id、格式、工作表/页数信息。",
            inputSchema={"type": "object", "properties": {}, "additionalProperties": False},
        ),
        _read_document_tool(
            "规划阶段查看资料结构或封面/PDF/扫描件图像（unit 为单元编号或工作表名，view/mode=text|image|layout；text 默认每次读取 150 行）。只需确认章节归属与分组即可调用 submit_book_plan，不要逐行通读全部名单。"
        ),
        types.Tool(
            name="inspect_docx_outline",
            description="查看 Word 文档的章节标题、表格位置与内嵌图片块索引（用于精确拆分联合文档如积分表+对阵表）。",
            inputSchema={
                "type": "object",
                "properties": {"file_id": {"type": "string"}},
                "required": ["file_id"],
                "additionalProperties": False,
            },
        ),
        types.Tool(
            name="inspect_template",
            description="读取当前选定秩序册模板的全部章节结构、名单母版和版式说明。",
            inputSchema={"type": "object", "properties": {}, "additionalProperties": False},
        ),
        types.Tool(
            name="submit_book_plan",
            description="提交最终整册编排规划（BookBuildPlan），明确所有文件的章节归属、封面底图、名单范围与排除项、缺失/未识别报告。",
            inputSchema={
                "type": "object",
                "properties": {
                    "event_title": {"type": "string"},
                    "subtitle": {"type": "string"},
                    "footer_lines": {"type": "array", "items": {"type": "string"}},
                    "cover": {
                        "type": "object",
                        "properties": {
                            "mode": {"type": "string", "enum": ["image", "background", "generated"]},
                            "file_id": {"type": ["string", "null"]},
                        },
                    },
                    "back_cover": {
                        "type": "object",
                        "properties": {"file_id": {"type": ["string", "null"]}},
                    },
                    "watermark": {
                        "type": "object",
                        "properties": {"file_id": {"type": ["string", "null"]}},
                    },
                    "roster_scope": {
                        "type": "object",
                        "properties": {
                            "instruction": {"type": "string"},
                            "groups": {"type": "array", "items": {"type": "string"}},
                            "sources": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "file_id": {"type": "string"},
                                        "included_sheets": {"type": "array", "items": {"type": "string"}},
                                        "excluded_sheets": {"type": "array", "items": {"type": "string"}},
                                        "group_map": {"type": "object"},
                                    },
                                    "required": ["file_id"],
                                },
                            },
                        },
                    },
                    "chapters": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "id": {"type": "string"},
                                "title": {"type": "string"},
                                "kind": {"type": "string"},
                                "toc_visible": {"type": "boolean"},
                                "landscape": {"type": ["boolean", "string"]},
                                "split_by_group_in_toc": {"type": "boolean"},
                                "sources": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "file_id": {"type": "string"},
                                            "sheets": {"type": "array", "items": {"type": "string"}},
                                            "start_pattern": {"type": "string"},
                                            "end_pattern": {"type": "string"},
                                            "caption": {"type": "string"},
                                        },
                                        "required": ["file_id"],
                                    },
                                },
                            },
                            "required": ["id", "title", "kind"],
                        },
                    },
                    "excluded_items": {"type": "array", "items": {"type": "string"}},
                    "missing_items": {"type": "array", "items": {"type": "string"}},
                    "unrecognized_files": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {"file": {"type": "string"}, "reason": {"type": "string"}},
                        },
                    },
                },
                "required": ["event_title", "roster_scope", "chapters", "missing_items"],
                "additionalProperties": True,
            },
        ),
    ]


def _draw_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="list_documents",
            description="列出上传的抽签结果文件及其格式、工作表/页数。",
            inputSchema={"type": "object", "properties": {}, "additionalProperties": False},
        ),
        _read_document_tool(
            "读取抽签结果文件（支持 Excel / Word / PDF / 图片；unit 为从 1 开始的单元编号或工作表名，view/mode=text|image|layout）。"
        ),
        types.Tool(
            name="inspect_target_book",
            description="查看目标秩序册 Word 中的所有赛程/分组表格及候选签位单元格/段落（含行上下文与胜负/名次警告），支持按 table_index 或 row_start 分页读取。",
            inputSchema={
                "type": "object",
                "properties": {
                    "table_index": {"type": "integer"},
                    "row_start": {"type": "integer"},
                    "row_limit": {"type": "integer"},
                },
                "additionalProperties": False,
            },
        ),
        types.Tool(
            name="submit_draw_patches",
            description="提交抽签后赛程替换补丁清单（仅替换抽签签位占位，严禁破坏胜负引用和名次引用）。",
            inputSchema={
                "type": "object",
                "properties": {
                    "summary": {"type": "string"},
                    "preserved_examples": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "location": {"type": "string"},
                                "text": {"type": "string"},
                                "reason": {"type": "string"},
                            },
                        },
                    },
                    "patches": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "target": {"type": "string", "enum": ["cell", "paragraph"]},
                                "table": {"type": "integer"},
                                "row": {"type": "integer"},
                                "col": {"type": "integer"},
                                "paragraph": {"type": "integer"},
                                "find": {"type": "string"},
                                "replace": {"type": "string"},
                                "reason": {"type": "string"},
                            },
                            "required": ["target", "find", "replace"],
                        },
                    },
                },
                "required": ["patches", "summary"],
                "additionalProperties": False,
            },
        ),
    ]


def _inspect_target_book_page(args: dict[str, Any]) -> dict[str, Any]:
    data = storage.read_json(storage.task_dir(TASK_ID) / "book_inspect.json") or {}
    t_idx = args.get("table_index")
    if t_idx is None:
        return {
            "file_name": data.get("file_name"),
            "table_count": data.get("table_count", 0),
            "tables_summary": [
                {
                    "table_index": t["table_index"],
                    "heading": t.get("heading"),
                    "row_count": t["row_count"],
                    "col_count": t["col_count"],
                    "headers": t.get("headers"),
                    "has_candidate_slot": t.get("has_candidate_slot"),
                }
                for t in (data.get("tables") or [])
            ],
            "candidate_paragraphs": data.get("candidate_paragraphs") or [],
        }
    tables = {t["table_index"]: t for t in (data.get("tables") or [])}
    tbl = tables.get(int(t_idx))
    if not tbl:
        return {"error": f"Unknown table_index: {t_idx}"}
    r_start = int(args.get("row_start") or 0)
    r_limit = min(int(args.get("row_limit") or 60), 100)
    rows = [r for r in (tbl.get("rows") or []) if r["row"] >= r_start][:r_limit]
    return {
        "table_index": tbl["table_index"],
        "heading": tbl.get("heading"),
        "row_count": tbl["row_count"],
        "col_count": tbl["col_count"],
        "headers": tbl.get("headers"),
        "rows": rows,
        "has_more": len([r for r in (tbl.get("rows") or []) if r["row"] >= r_start]) > r_limit,
    }


def _normalize_read_args(args: dict[str, Any]) -> dict[str, Any]:
    file_id = str(args.get("file_id") or "")
    unit_raw = args.get("unit")
    unit_idx: Any = 1 if unit_raw is None else unit_raw
    file_info: dict[str, Any] | None = None
    try:
        file_info = document_bank.by_id(TASK_ID, file_id) if file_id else None
    except document_bank.DocumentError:
        file_info = None
    units = (file_info or {}).get("units") or []
    if isinstance(unit_raw, str):
        stripped = unit_raw.strip()
        if stripped.isdigit():
            unit_idx = int(stripped)
        elif stripped:
            matched = next(
                (u["index"] for u in units if str(u.get("label") or "").strip() == stripped),
                None,
            )
            if matched is None:
                matched = next(
                    (u["index"] for u in units if stripped in str(u.get("label") or "")),
                    None,
                )
            if matched is None:
                available = ", ".join(f"{u.get('index')}:{u.get('label')}" for u in units)
                raise document_bank.DocumentError(f"未找到读取单元 '{stripped}'。可用单元: {available}")
            unit_idx = matched
        else:
            unit_idx = 1
    view = str(args.get("view") or args.get("mode") or "text")
    out: dict[str, Any] = {
        "file_id": file_id,
        "unit": unit_idx,
        "view": view,
    }
    if args.get("start") is not None:
        out["start"] = args["start"]
    if args.get("limit") is not None:
        out["limit"] = args["limit"]
    if MODE == "plan" and view == "text":
        unit_size = next(
            (int(u.get("size") or 0) for u in units if u.get("index") == unit_idx),
            0,
        )
        if out.get("limit") is None:
            out["limit"] = 150
        elif unit_size > 40:
            try:
                out["limit"] = min(200, max(int(out["limit"]), 80))
            except (TypeError, ValueError):
                out["limit"] = 150
    return out


def _dispatch_tool(name: str, args: dict[str, Any]) -> types.CallToolResult:
    _trace("tool_call", name, args=args)
    try:
        if name == "list_documents":
            payload = document_bank.list_documents(TASK_ID)
            return _text_result(json.dumps(payload, ensure_ascii=False))

        if name == "read_document":
            norm = _normalize_read_args(args)
            if MODE == "plan":
                start_val = int(norm.get("start") or 1)
                limit_val = int(norm.get("limit") or 150)
                read_count, overlapped = _record_plan_window(
                    norm["file_id"], norm["unit"], norm["view"], start_val, limit_val
                )
                if read_count > 5:
                    raise document_bank.DocumentError(
                        f"规划阶段已在 {norm['file_id']} 单元 {norm['unit']} 连续读取 {read_count - 1} 次。"
                        "你已充分了解文件结构与表头，请停止重复翻页，立即调用 submit_book_plan 提交整册规划。"
                    )
            else:
                read_count, overlapped = 1, False
            payload, image = document_bank.read_document(TASK_ID, norm)
            _mark_plan_touched(norm["file_id"])
            if MODE == "plan" and (read_count >= 2 or overlapped) and isinstance(payload, dict):
                payload = dict(payload)
                payload["plan_hint"] = (
                    f"规划阶段只需确认文件结构与分组，无需逐行通读全文（该单元已读 {read_count} 次）；"
                    "结构明确后请直接调用 submit_book_plan。"
                )
            content: list[types.ContentBlock] = [
                types.TextContent(type="text", text=json.dumps(payload, ensure_ascii=False))
            ]
            if image:
                content.append(
                    types.ImageContent(
                        type="image",
                        data=base64.b64encode(image).decode("ascii"),
                        mimeType="image/jpeg",
                    )
                )
            return types.CallToolResult(isError=False, content=content)

        if MODE == "plan":
            if name == "inspect_docx_outline":
                fid = str(args.get("file_id") or "")
                res = _inspect_docx_outline_impl(fid)
                if not res.get("error"):
                    _mark_plan_touched(fid)
                return _text_result(json.dumps(res, ensure_ascii=False), error=bool(res.get("error")))
            if name == "inspect_template":
                task = storage.read_task(TASK_ID)
                return _text_result(json.dumps(task.get("template_spec") or {}, ensure_ascii=False))
            if name == "submit_book_plan":
                errors = _validate_book_plan(args)
                if errors:
                    return _text_result(json.dumps({"errors": errors}, ensure_ascii=False), error=True)
                storage.write_json(storage.task_dir(TASK_ID) / "book_plan.json", args)
                _trace("submitted", "submit_book_plan", chapters=len(args.get("chapters") or []))
                return _text_result(json.dumps({"ok": True}, ensure_ascii=False))

        if MODE == "draw":
            if name == "inspect_target_book":
                res = _inspect_target_book_page(args)
                return _text_result(json.dumps(res, ensure_ascii=False), error=bool(res.get("error")))
            if name == "submit_draw_patches":
                errors, cleaned = _validate_draw_patches(args)
                if errors:
                    return _text_result(json.dumps({"errors": errors}, ensure_ascii=False), error=True)
                storage.write_json(storage.task_dir(TASK_ID) / "draw_patches.json", cleaned)
                _trace("submitted", "submit_draw_patches", count=len(cleaned["patches"]))
                return _text_result(json.dumps({"ok": True, "count": len(cleaned["patches"])}, ensure_ascii=False))

        return _text_result(json.dumps({"error": f"Unknown tool: {name}"}, ensure_ascii=False), error=True)
    except (document_bank.DocumentError, OSError, ValueError) as exc:
        msg = str(exc).strip() or "工具执行失败。"
        _trace("tool_error", name, error=msg)
        return _text_result(json.dumps({"error": msg}, ensure_ascii=False), error=True)


async def _serve() -> None:
    async def on_list_tools(_context: Any, _params: Any) -> types.ListToolsResult:
        tools = _plan_tools() if MODE == "plan" else _draw_tools()
        return types.ListToolsResult(tools=tools)

    async def on_call_tool(_context: Any, params: Any) -> types.CallToolResult:
        return _dispatch_tool(params.name, dict(params.arguments or {}))

    server = Server(
        "book" if MODE == "plan" else "draw",
        on_list_tools=on_list_tools,
        on_call_tool=on_call_tool,
    )
    async with stdio_server() as (reader, writer):
        await server.run(reader, writer, server.create_initialization_options())


def main() -> int:
    global TASK_ID, MODE
    parser = argparse.ArgumentParser()
    parser.add_argument("--task", required=True)
    parser.add_argument("--mode", choices=("plan", "draw"), required=True)
    ns = parser.parse_args()
    TASK_ID = ns.task
    MODE = ns.mode
    try:
        asyncio.run(_serve())
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
