from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from documents import TEXTUTIL, find_soffice

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _run_soffice_convert(
    src_path: Path,
    out_dir: Path,
    target_ext: str,
    timeout: float = 120.0,
) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "未检测到 LibreOffice (soffice)，无法转换或导出 PDF。请安装 LibreOffice 或设置 LIBREOFFICE_BIN。"
        )
    src_path = Path(src_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="lo-prof-") as prof_dir:
        cmd = [
            soffice,
            f"-env:UserInstallation=file://{prof_dir}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--convert-to",
            target_ext,
            "--outdir",
            str(out_dir),
            str(src_path),
        ]
        res = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
        expected = out_dir / f"{src_path.stem}.{target_ext}"
        if res.returncode != 0 or not expected.is_file():
            detail = (res.stderr or res.stdout or "").strip()
            raise RuntimeError(
                f"LibreOffice 转换 {src_path.name} -> .{target_ext} 失败 (exit={res.returncode}): {detail}"
            )
        return expected


def export_pdf(docx_path: Path, pdf_path: Path | None = None, timeout: float = 120.0) -> Path:
    """Export a Word document (.docx/.doc) to PDF via headless LibreOffice."""
    docx_path = Path(docx_path).resolve()
    target = Path(pdf_path).resolve() if pdf_path else docx_path.with_suffix(".pdf")
    target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="lo-pdf-") as tmp_dir:
        produced = _run_soffice_convert(docx_path, Path(tmp_dir), "pdf", timeout=timeout)
        tmp_target = target.with_suffix(".pdf.tmp")
        shutil.copy2(produced, tmp_target)
        tmp_target.replace(target)
    return target


def convert_to_docx(src_path: Path, dst_path: Path | None = None, timeout: float = 120.0) -> Path:
    """Convert .doc or .wps (or normalize .docx) into a clean .docx file."""
    src_path = Path(src_path).resolve()
    suffix = src_path.suffix.lower()
    target = Path(dst_path).resolve() if dst_path else src_path.with_suffix(".docx")
    target.parent.mkdir(parents=True, exist_ok=True)
    if suffix == ".docx":
        if src_path != target:
            shutil.copy2(src_path, target)
        return target
    with tempfile.TemporaryDirectory(prefix="lo-docx-") as tmp_dir:
        tmp_out = Path(tmp_dir)
        errors: list[str] = []
        if find_soffice():
            try:
                produced = _run_soffice_convert(src_path, tmp_out, "docx", timeout=timeout)
                tmp_target = target.with_suffix(".docx.tmp")
                shutil.copy2(produced, tmp_target)
                tmp_target.replace(target)
                return target
            except Exception as exc:
                errors.append(str(exc))
        if suffix in {".doc", ".wps"} and TEXTUTIL.is_file():
            produced = tmp_out / f"{src_path.stem}.docx"
            res = subprocess.run(
                [str(TEXTUTIL), "-convert", "docx", "-output", str(produced), str(src_path)],
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
            if res.returncode == 0 and produced.is_file():
                tmp_target = target.with_suffix(".docx.tmp")
                shutil.copy2(produced, tmp_target)
                tmp_target.replace(target)
                return target
            errors.append((res.stderr or res.stdout or "").strip())
        raise RuntimeError(f"无法将 {src_path.name} 转换为 .docx: {'; '.join(errors)}")


def _normalize_heading_text(text: str) -> str:
    return re.sub(r"\s+", "", text or "").strip()


def _detect_footer_page_number(page: pymupdf.Page) -> int | None:
    """Try to parse a printed Arabic page number from the bottom footer strip of a PDF page."""
    rect = page.rect
    y_top = max(rect.height * 0.90, rect.height - 68.0)
    footer_clip = pymupdf.Rect(rect.x0, y_top, rect.x1, rect.height)
    text = (page.get_text("text", clip=footer_clip) or "").strip()
    if not text:
        return None
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    # Prefer decorated footer page marks like "— 1 —" or "第 1 页" first
    for line in reversed(lines):
        m = re.fullmatch(r"(?:[-—–·]+\s*([0-9]{1,4})\s*[-—–·]+|第\s*([0-9]{1,4})\s*页)", line)
        if m:
            return int(m.group(1) or m.group(2))
    for line in reversed(lines):
        m = re.fullmatch(r"([0-9]{1,4})", line)
        if m:
            return int(m.group(1))
    return None


def _section_page_restarts(doc: Document) -> list[dict]:
    """Return metadata per section: whether it has footer page numbers and pgNumType w:start."""
    info: list[dict] = []
    for idx, section in enumerate(doc.sections):
        sect_pr = section._sectPr
        pg_num = sect_pr.find(qn("w:pgNumType"))
        start_val = None
        if pg_num is not None:
            raw_start = pg_num.get(qn("w:start"))
            if raw_start and raw_start.isdigit():
                start_val = int(raw_start)
        has_footer_page = False
        for footer_attr in ("footer", "first_page_footer", "even_page_footer"):
            footer = getattr(section, footer_attr, None)
            if footer is None:
                continue
            xml = footer._element.xml
            if "PAGE" in xml:
                has_footer_page = True
                break
        info.append(
            {
                "section_index": idx,
                "start": start_val,
                "has_footer_page": has_footer_page,
            }
        )
    return info


def _find_bookmarked_paragraphs(doc: Document) -> dict[str, tuple[str, int]]:
    """Return bookmark_name -> (paragraph_text, section_index) by walking body elements directly."""
    found: dict[str, tuple[str, int]] = {}
    sec_idx = 0
    for child in doc.element.body:
        if child.tag == qn("w:p"):
            txt = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            for bm in child.findall(qn("w:bookmarkStart")):
                name = bm.get(qn("w:name")) or ""
                if name:
                    found[name] = (txt, sec_idx)
            p_pr = child.find(qn("w:pPr"))
            if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
                sec_idx += 1
        elif child.tag == qn("w:tbl"):
            for p_el in child.iter(qn("w:p")):
                txt = "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()
                for bm in p_el.findall(qn("w:bookmarkStart")):
                    name = bm.get(qn("w:name")) or ""
                    if name:
                        found[name] = (txt, sec_idx)
    return found


def _update_pageref_fields(doc: Document, page_by_anchor: dict[str, int]) -> int:
    """Update cached <w:t> inside PAGEREF <anchor> fields across paragraphs and SDTs."""
    updated = 0
    for p in doc._element.body.iter(qn("w:p")):
        # Walk direct and hyperlink-nested runs in order
        elements: list = []
        for child in p:
            if child.tag in {qn("w:r"), qn("w:fldSimple")}:
                elements.append(child)
            elif child.tag == qn("w:hyperlink"):
                for sub in child:
                    if sub.tag in {qn("w:r"), qn("w:fldSimple")}:
                        elements.append(sub)
        active_anchor: str | None = None
        in_display = False
        first_display_written = False
        instr_buf = ""
        in_instr = False
        for node in elements:
            if node.tag == qn("w:fldSimple"):
                instr = node.get(qn("w:instr")) or ""
                m = re.search(r"PAGEREF\s+([^\s\\]+)", instr)
                if m and m.group(1) in page_by_anchor:
                    for t in node.iter(qn("w:t")):
                        t.text = str(page_by_anchor[m.group(1)])
                        updated += 1
                        break
                continue
            fld_char = node.find(qn("w:fldChar"))
            if fld_char is not None:
                ftype = fld_char.get(qn("w:fldCharType"))
                if ftype == "begin":
                    active_anchor = None
                    in_display = False
                    first_display_written = False
                    instr_buf = ""
                    in_instr = True
                elif ftype == "separate":
                    in_instr = False
                    m = re.search(r"PAGEREF\s+([^\s\\]+)", instr_buf)
                    if m and m.group(1) in page_by_anchor:
                        active_anchor = m.group(1)
                        in_display = True
                elif ftype == "end":
                    active_anchor = None
                    in_display = False
                    in_instr = False
                continue
            if in_instr:
                for it in node.findall(qn("w:instrText")):
                    instr_buf += it.text or ""
            elif in_display and active_anchor is not None:
                t_elems = node.findall(qn("w:t"))
                if t_elems:
                    if not first_display_written:
                        t_elems[0].text = str(page_by_anchor[active_anchor])
                        first_display_written = True
                        updated += 1
                    else:
                        t_elems[0].text = ""
    return updated


def _update_sdt_toc_entries(doc: Document, heading_pages: list[tuple[str, int]]) -> int:
    """Fallback updater for Word <w:sdt> TOC blocks using cached PAGEREF/tab trailing runs."""
    updated = 0
    normalized_lookup = {
        _normalize_heading_text(title): page for title, page in heading_pages if title
    }
    for sdt in doc._element.body.iter(qn("w:sdt")):
        for p in sdt.iter(qn("w:p")):
            full = "".join(t.text or "" for t in p.iter(qn("w:t")))
            # Strip trailing digits/tabs
            stripped = re.sub(r"[\t\s0-9]+$", "", full)
            norm = _normalize_heading_text(stripped)
            target_page = normalized_lookup.get(norm)
            if target_page is None:
                for cand_norm, cand_page in normalized_lookup.items():
                    if cand_norm and (cand_norm in norm or norm in cand_norm):
                        target_page = cand_page
                        break
            if target_page is None:
                continue
            # Replace the last numeric run if present
            runs = list(p.iter(qn("w:r")))
            for run in reversed(runs):
                t = run.find(qn("w:t"))
                if t is not None and re.fullmatch(r"\s*\d{1,4}\s*", t.text or ""):
                    t.text = str(target_page)
                    updated += 1
                    break
    return updated


def enable_update_fields_on_open(doc: Document) -> None:
    """Set <w:updateFields w:val="true"/> in settings.xml so Word refreshes fields cleanly."""
    settings_el = doc.settings._element
    existing = settings_el.find(qn("w:updateFields"))
    if existing is not None:
        existing.set(qn("w:val"), "true")
    else:
        el = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
        settings_el.append(el)


def _match_pdf_toc_to_targets(
    pdf_doc: pymupdf.Document,
    doc: Document,
    toc_targets: list[dict] | None,
) -> tuple[dict[str, int], list[tuple[str, int]]]:
    """Compute printed page number for every TOC bookmark/heading from a Pass-1 PDF."""
    pdf_toc = pdf_doc.get_toc(simple=True)  # [(lvl, title, pdf_page_1based), ...]
    sec_info = _section_page_restarts(doc)
    bookmarks = _find_bookmarked_paragraphs(doc)

    # Build ordered list of body headings from toc_targets or bookmarks
    ordered_targets: list[tuple[str, str, int]] = []  # (anchor, title, sec_idx)
    if toc_targets:
        for item in toc_targets:
            anchor = str(item.get("anchor") or "")
            title = str(item.get("title") or "")
            sec_idx = bookmarks.get(anchor, (title, 1))[1]
            ordered_targets.append((anchor, title, sec_idx))
    else:
        for anchor, (title, sec_idx) in bookmarks.items():
            if anchor.startswith("_toc_"):
                ordered_targets.append((anchor, title, sec_idx))

    # Map PDF TOC entries sequentially / by normalized title
    pdf_entries_norm = [(_normalize_heading_text(t), int(pg), t) for _lvl, t, pg in pdf_toc]

    # Match each target heading to its 1-based physical PDF page
    matched_pdf_pages: list[tuple[str, str, int, int]] = []  # (anchor, title, sec_idx, pdf_page)
    cursor = 0
    for anchor, title, sec_idx in ordered_targets:
        norm = _normalize_heading_text(title)
        found_pdf_page: int | None = None
        for idx in range(cursor, len(pdf_entries_norm)):
            cand_norm, cand_pg, _ = pdf_entries_norm[idx]
            if cand_norm == norm or (norm and cand_norm and (norm in cand_norm or cand_norm in norm)):
                found_pdf_page = cand_pg
                cursor = idx + 1
                break
        if found_pdf_page is None:
            for idx, (cand_norm, cand_pg, _) in enumerate(pdf_entries_norm):
                if cand_norm == norm:
                    found_pdf_page = cand_pg
                    break
        if found_pdf_page is None and title.strip():
            # Fallback: search visible text across PDF pages starting after page 1
            for page_idx in range(pdf_doc.page_count):
                txt = _normalize_heading_text(pdf_doc[page_idx].get_text("text"))
                if norm and norm in txt:
                    found_pdf_page = page_idx + 1
                    break
        if found_pdf_page is not None:
            matched_pdf_pages.append((anchor, title, sec_idx, found_pdf_page))

    # Determine page offset relative to section restart (<w:pgNumType w:start="1"/>)
    # Find the earliest section with explicit start != None (defaulting to body sec where targets begin)
    restart_sections = {s["section_index"]: s["start"] for s in sec_info if s["start"] is not None}
    offset_by_sec: dict[int, int] = {}
    if matched_pdf_pages:
        first_target_pdf_page = matched_pdf_pages[0][3]
        first_target_sec = matched_pdf_pages[0][2]
        # Prefer actual printed footer number on first_target_pdf_page if readable
        footer_num = _detect_footer_page_number(pdf_doc[first_target_pdf_page - 1])
        if footer_num is not None:
            base_offset = first_target_pdf_page - footer_num
        elif first_target_sec in restart_sections:
            base_offset = first_target_pdf_page - int(restart_sections[first_target_sec])
        elif restart_sections:
            earliest_sec = min(restart_sections)
            # Check if first target is in or after earliest_sec
            if first_target_sec >= earliest_sec:
                base_offset = first_target_pdf_page - int(restart_sections[earliest_sec])
            else:
                base_offset = 0
        else:
            base_offset = 0

        for sec_s in sec_info:
            si = sec_s["section_index"]
            if si in restart_sections and si != first_target_sec:
                # Find first matched heading in this restarted section
                in_sec = [m for m in matched_pdf_pages if m[2] == si]
                if in_sec:
                    f_num = _detect_footer_page_number(pdf_doc[in_sec[0][3] - 1])
                    if f_num is not None:
                        offset_by_sec[si] = in_sec[0][3] - f_num
                    else:
                        offset_by_sec[si] = in_sec[0][3] - int(restart_sections[si])
                    continue
            offset_by_sec[si] = base_offset

    page_by_anchor: dict[str, int] = {}
    heading_pages: list[tuple[str, int]] = []
    for anchor, title, sec_idx, pdf_page in matched_pdf_pages:
        calc_printed = max(1, pdf_page - offset_by_sec.get(sec_idx, 0))
        footer_num = _detect_footer_page_number(pdf_doc[pdf_page - 1])
        printed = calc_printed if (sec_idx in offset_by_sec or footer_num is None) else footer_num
        if anchor:
            page_by_anchor[anchor] = printed
        heading_pages.append((title, printed))

    # Also include raw PDF TOC items for SDT fallback
    if not heading_pages and pdf_toc:
        for _lvl, title, pdf_page in pdf_toc:
            footer_num = _detect_footer_page_number(pdf_doc[pdf_page - 1])
            heading_pages.append((title, footer_num or pdf_page))

    return page_by_anchor, heading_pages


def sync_toc_and_export_pdf(
    docx_path: Path,
    pdf_path: Path | None = None,
    toc_targets: list[dict] | None = None,
    timeout: float = 120.0,
) -> tuple[Path, list[dict]]:
    """Run Pass-1 PDF export, compute exact body page numbers for TOC, write back to DOCX, and run Pass-2 PDF export."""
    docx_path = Path(docx_path).resolve()
    target_pdf = Path(pdf_path).resolve() if pdf_path else docx_path.with_suffix(".pdf")

    with tempfile.TemporaryDirectory(prefix="toc-sync-") as tmp_dir:
        pass1_pdf = Path(tmp_dir) / "pass1.pdf"
        export_pdf(docx_path, pass1_pdf, timeout=timeout)

        doc = Document(str(docx_path))
        enable_update_fields_on_open(doc)

        with pymupdf.open(pass1_pdf) as pdf_doc:
            page_by_anchor, heading_pages = _match_pdf_toc_to_targets(pdf_doc, doc, toc_targets)

        changed = 0
        if page_by_anchor:
            changed += _update_pageref_fields(doc, page_by_anchor)
        if heading_pages:
            changed += _update_sdt_toc_entries(doc, heading_pages)

        if changed or toc_targets is not None:
            tmp_docx = docx_path.with_suffix(".docx.tmp")
            doc.save(str(tmp_docx))
            tmp_docx.replace(docx_path)

    # Pass 2 export so PDF TOC displays identical updated page numbers
    export_pdf(docx_path, target_pdf, timeout=timeout)

    resolved_entries: list[dict] = []
    if toc_targets:
        for item in toc_targets:
            anchor = str(item.get("anchor") or "")
            entry = dict(item)
            if anchor in page_by_anchor:
                entry["page"] = page_by_anchor[anchor]
            resolved_entries.append(entry)
    else:
        for title, page in heading_pages:
            resolved_entries.append({"title": title, "page": page})
    return target_pdf, resolved_entries


def render_pdf_pages(
    pdf_path: Path,
    out_dir: Path,
    zoom: float = 1.5,
    max_pages: int = 100,
) -> list[Path]:
    """Render pages of a PDF to PNG images for visual inspection."""
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    rendered: list[Path] = []
    with pymupdf.open(pdf_path) as pdf_doc:
        count = min(pdf_doc.page_count, max_pages)
        matrix = pymupdf.Matrix(zoom, zoom)
        for idx in range(count):
            pix = pdf_doc[idx].get_pixmap(matrix=matrix)
            out_file = out_dir / f"page_{idx + 1:03d}.png"
            pix.save(str(out_file))
            rendered.append(out_file)
    return rendered
