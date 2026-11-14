from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn

from book_pdf import sync_toc_and_export_pdf

# Patterns commonly used in pre-draw schedules (draw lots) vs knockout win/loss or rank references
DRAW_SLOT_RE = re.compile(
    r"(?:^|[^0-9A-Za-z])"
    r"(?:[\u4e00-\u9fff]{1,4})?[A-Za-z]\s*\d{1,2}"
    r"(?=$|[^0-9A-Za-z])"
)
WIN_LOSS_RE = re.compile(r"(?:\d+\s*[胜负]|第\s*\d+\s*场[胜负]|胜者|败者|第[一二三四五六七八\d]+名)")
KNOCKOUT_STAGE_RE = re.compile(r"淘汰赛|交叉赛|半决赛|决赛|名次赛|排位赛|季军|三四名|八强|四强|1/4|1/2")


def _paragraph_full_text(p_el) -> str:
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


def _cell_full_text(cell) -> str:
    return "\n".join(_paragraph_full_text(p._p) for p in cell.paragraphs).strip()


def _replace_span_in_paragraph_xml(p_el, start: int, end: int, replacement: str) -> bool:
    """Replace character range [start:end) across <w:t> elements in p_el while preserving first run formatting."""
    segments: list[tuple[Any, int, int]] = []
    pos = 0
    for t_el in p_el.iter(qn("w:t")):
        txt = t_el.text or ""
        ln = len(txt)
        segments.append((t_el, pos, pos + ln))
        pos += ln
    if start < 0 or end > pos or start >= end:
        return False

    wrote = False
    for t_el, seg_start, seg_end in segments:
        if seg_end <= start or seg_start >= end:
            continue
        old = t_el.text or ""
        local_s = max(0, start - seg_start)
        local_e = min(len(old), end - seg_start)
        if not wrote:
            t_el.text = old[:local_s] + replacement + old[local_e:]
            # Preserve surrounding spaces cleanly if leading/trailing whitespace introduced
            if t_el.text.startswith(" ") or t_el.text.endswith(" "):
                t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            wrote = True
        else:
            t_el.text = old[:local_s] + old[local_e:]
    return wrote


def replace_text_preserve_format(container_paragraphs: list, find_text: str, replace_text: str) -> int:
    """Replace all occurrences of find_text in a list of python-docx Paragraph objects preserving run styles."""
    if not find_text:
        return 0
    count = 0
    for p in container_paragraphs:
        while True:
            full = _paragraph_full_text(p._p)
            idx = full.find(find_text)
            if idx < 0:
                break
            if not _replace_span_in_paragraph_xml(p._p, idx, idx + len(find_text), replace_text):
                break
            count += 1
    return count


def inspect_target_book(docx_path: Path) -> dict[str, Any]:
    """Extract structured tables & context from an existing Word order book for safe draw-slot planning."""
    doc = Document(str(docx_path))

    # Walk body children sequentially; top-level <w:p> and <w:tbl> order matches doc.paragraphs / doc.tables
    last_heading = ""
    tables_info: list[dict[str, Any]] = []
    candidate_paragraphs: list[dict[str, Any]] = []
    p_cursor = 0
    t_cursor = 0

    for child in doc.element.body:
        if child.tag == qn("w:p"):
            cur_p_idx = p_cursor
            p_cursor += 1
            txt = _paragraph_full_text(child).strip()
            if txt and len(txt) <= 45 and (
                re.match(r"^[一二三四五六七八九十百]+[、.．]", txt)
                or any(k in txt for k in ("赛程", "日程", "分组", "对阵", "名单"))
            ):
                last_heading = txt
            if txt and (DRAW_SLOT_RE.search(txt) or WIN_LOSS_RE.search(txt)):
                candidate_paragraphs.append(
                    {
                        "paragraph_index": cur_p_idx,
                        "heading": last_heading,
                        "text": txt,
                        "has_win_loss_or_rank": bool(WIN_LOSS_RE.search(txt)),
                    }
                )
        elif child.tag == qn("w:tbl"):
            t_idx = t_cursor
            t_cursor += 1
            if t_idx >= len(doc.tables):
                continue
            tbl = doc.tables[t_idx]
            rows_dump: list[dict[str, Any]] = []
            has_candidate_slot = False
            headers = [_cell_full_text(c) for c in tbl.rows[0].cells] if tbl.rows else []
            for r_idx, row in enumerate(tbl.rows):
                cells_dump: list[dict[str, Any]] = []
                # Deduplicate horizontal merged cell objects cleanly by tracking tc element identity
                seen_tcs: set[int] = set()
                row_has_interest = False
                for c_idx, cell in enumerate(row.cells):
                    tc_id = id(cell._tc)
                    if tc_id in seen_tcs:
                        continue
                    seen_tcs.add(tc_id)
                    txt = _cell_full_text(cell)
                    is_slot = bool(DRAW_SLOT_RE.search(txt))
                    is_wl = bool(WIN_LOSS_RE.search(txt))
                    if is_slot or is_wl:
                        row_has_interest = True
                        if is_slot:
                            has_candidate_slot = True
                    cells_dump.append(
                        {
                            "col": c_idx,
                            "text": txt,
                            "has_slot_pattern": is_slot,
                            "has_win_loss_or_rank": is_wl,
                        }
                    )
                if r_idx == 0 or row_has_interest or len(tbl.rows) <= 25:
                    rows_dump.append({"row": r_idx, "cells": cells_dump})
            tables_info.append(
                {
                    "table_index": t_idx,
                    "heading": last_heading,
                    "row_count": len(tbl.rows),
                    "col_count": len(tbl.columns),
                    "headers": headers,
                    "has_candidate_slot": has_candidate_slot,
                    "rows": rows_dump,
                }
            )

    return {
        "file_name": Path(docx_path).name,
        "table_count": len(doc.tables),
        "paragraph_count": len(doc.paragraphs),
        "tables": tables_info,
        "candidate_paragraphs": candidate_paragraphs[:120],
    }


def apply_draw_patches(
    target_docx: Path,
    patches: list[dict[str, Any]],
    out_docx: Path,
    out_pdf: Path | None = None,
    *,
    preserved_examples: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Apply exact cell/paragraph replacements to target_docx preserving formatting, then sync TOC & export PDF."""
    target_docx = Path(target_docx).resolve()
    out_docx = Path(out_docx).resolve()
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(target_docx, out_docx)

    doc = Document(str(out_docx))
    applied: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []

    for patch in patches:
        tgt_kind = str(patch.get("target") or "cell")
        find_text = str(patch.get("find") or "")
        replace_text = str(patch.get("replace") or "")
        if not find_text or find_text == replace_text:
            continue

        if tgt_kind == "cell":
            t_idx = int(patch.get("table", -1))
            r_idx = int(patch.get("row", -1))
            c_idx = int(patch.get("col", -1))
            if not (0 <= t_idx < len(doc.tables)):
                skipped.append({**patch, "reason": "table index out of range"})
                continue
            tbl = doc.tables[t_idx]
            if not (0 <= r_idx < len(tbl.rows)) or not (0 <= c_idx < len(tbl.rows[r_idx].cells)):
                skipped.append({**patch, "reason": "row/col index out of range"})
                continue
            cell = tbl.rows[r_idx].cells[c_idx]
            before = _cell_full_text(cell)
            if find_text not in before:
                skipped.append({**patch, "reason": f"'{find_text}' not found in cell text '{before}'"})
                continue
            # Extra safety: never touch explicit win/loss references unless explicitly forced
            if WIN_LOSS_RE.search(find_text):
                skipped.append({**patch, "reason": "refused to replace win/loss/ranking token"})
                continue
            replaced_hits = replace_text_preserve_format(cell.paragraphs, find_text, replace_text)
            after = _cell_full_text(cell)
            applied.append(
                {
                    "location": f"table[{t_idx}].row[{r_idx}].col[{c_idx}]",
                    "before": before,
                    "after": after,
                    "find": find_text,
                    "replace": replace_text,
                    "hits": replaced_hits,
                    "reason": str(patch.get("reason") or ""),
                }
            )
        elif tgt_kind == "paragraph":
            p_idx = int(patch.get("paragraph", -1))
            if not (0 <= p_idx < len(doc.paragraphs)):
                skipped.append({**patch, "reason": "paragraph index out of range"})
                continue
            para = doc.paragraphs[p_idx]
            before = para.text
            if find_text not in before:
                skipped.append({**patch, "reason": f"'{find_text}' not found in paragraph"})
                continue
            if WIN_LOSS_RE.search(find_text):
                skipped.append({**patch, "reason": "refused to replace win/loss/ranking token"})
                continue
            hits = replace_text_preserve_format([para], find_text, replace_text)
            applied.append(
                {
                    "location": f"paragraph[{p_idx}]",
                    "before": before,
                    "after": para.text,
                    "find": find_text,
                    "replace": replace_text,
                    "hits": hits,
                    "reason": str(patch.get("reason") or ""),
                }
            )

    # Scan remaining win/loss or ranking tokens left preserved across tables for transparent audit report
    preserved_found: list[dict[str, Any]] = list(preserved_examples or [])
    if not preserved_found:
        seen_pres: set[str] = set()
        for t_idx, tbl in enumerate(doc.tables):
            for r_idx, row in enumerate(tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    txt = _cell_full_text(cell)
                    if txt and WIN_LOSS_RE.search(txt) and txt not in seen_pres:
                        seen_pres.add(txt)
                        preserved_found.append(
                            {"location": f"table[{t_idx}].row[{r_idx}].col[{c_idx}]", "text": txt}
                        )
                        if len(preserved_found) >= 30:
                            break
                if len(preserved_found) >= 30:
                    break

    tmp_docx = out_docx.with_suffix(".docx.tmp")
    doc.save(str(tmp_docx))
    tmp_docx.replace(out_docx)

    target_pdf = out_pdf.resolve() if out_pdf else out_docx.with_suffix(".pdf")
    _, toc_entries = sync_toc_and_export_pdf(out_docx, target_pdf, timeout=180.0)

    return {
        "docx_path": str(out_docx),
        "pdf_path": str(target_pdf),
        "replaced_count": len(applied),
        "skipped_count": len(skipped),
        "preserved_count": len(preserved_found),
        "changes": applied,
        "skipped": skipped,
        "preserved": preserved_found,
        "toc": toc_entries,
    }
