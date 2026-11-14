from __future__ import annotations

import copy
import datetime as dt
import io
import re
import tempfile
from pathlib import Path
from typing import Any

import openpyxl
import pymupdf
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Emu, Mm, Pt
from PIL import Image

from book_pdf import _run_soffice_convert, convert_to_docx, enable_update_fields_on_open, sync_toc_and_export_pdf
from cover_art import generate_cover_from_background
from presets import PRESETS

EMU_PER_MM = 36000
TWIPS_PER_MM = 56.6929133858
EMU_PER_TWIP = 635
CN_NUMS = [
    "",
    "一",
    "二",
    "三",
    "四",
    "五",
    "六",
    "七",
    "八",
    "九",
    "十",
    "十一",
    "十二",
    "十三",
    "十四",
    "十五",
    "十六",
    "十七",
    "十八",
    "十九",
    "二十",
]


def _cn_index(n: int) -> str:
    if 1 <= n < len(CN_NUMS):
        return CN_NUMS[n]
    return str(n)


def _set_run_font(run, font_name: str, size_pt: float | None = None, *, bold: bool | None = None) -> None:
    if font_name:
        run.font.name = font_name
        r_pr = run._r.get_or_add_rPr()
        r_fonts = r_pr.find(qn("w:rFonts"))
        if r_fonts is None:
            r_fonts = OxmlElement("w:rFonts")
            r_pr.insert(0, r_fonts)
        for attr in ("ascii", "eastAsia", "hAnsi", "cs"):
            r_fonts.set(qn(f"w:{attr}"), font_name)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
        r_pr = run._r.get_or_add_rPr()
        for tag in ("w:b", "w:bCs"):
            for existing in r_pr.findall(qn(tag)):
                r_pr.remove(existing)
            el = OxmlElement(tag)
            if bold is False:
                el.set(qn("w:val"), "0")
            r_pr.append(el)


def _set_element_font(
    r_el,
    font_name: str,
    size_pt: float | None = None,
    *,
    bold: bool | None = None,
) -> None:
    r_pr = r_el.find(qn("w:rPr"))
    if r_pr is None:
        r_pr = OxmlElement("w:rPr")
        r_el.insert(0, r_pr)
    if font_name:
        r_fonts = r_pr.find(qn("w:rFonts"))
        if r_fonts is None:
            r_fonts = OxmlElement("w:rFonts")
            r_pr.insert(0, r_fonts)
        for attr in ("ascii", "eastAsia", "hAnsi", "cs"):
            r_fonts.set(qn(f"w:{attr}"), font_name)
    if size_pt is not None:
        val = str(int(round(size_pt * 2)))
        for tag in ("w:sz", "w:szCs"):
            sz = r_pr.find(qn(tag))
            if sz is None:
                sz = OxmlElement(tag)
                r_pr.append(sz)
            sz.set(qn("w:val"), val)
    if bold is not None:
        for tag in ("w:b", "w:bCs"):
            for existing in r_pr.findall(qn(tag)):
                r_pr.remove(existing)
            el = OxmlElement(tag)
            if bold is False:
                el.set(qn("w:val"), "0")
            r_pr.append(el)


def _configure_section(
    section,
    fmt: dict[str, Any],
    *,
    landscape: bool = False,
    zero_margins: bool = False,
) -> tuple[int, int]:
    """Configure section geometry, 22×28 page grid, and return printable (width_twips, height_twips)."""
    page_cfg = fmt.get("page", {})
    w_mm = float(page_cfg.get("width_mm", 210.0))
    h_mm = float(page_cfg.get("height_mm", 297.0))
    if landscape:
        w_mm, h_mm = max(w_mm, h_mm), min(w_mm, h_mm)
        section.orientation = WD_ORIENT.LANDSCAPE
        margins = page_cfg.get("landscape_margins_mm") or {"top": 20.0, "bottom": 20.0, "left": 22.0, "right": 22.0}
    else:
        w_mm, h_mm = min(w_mm, h_mm), max(w_mm, h_mm)
        section.orientation = WD_ORIENT.PORTRAIT
        margins = page_cfg.get("margins_mm") or {"top": 30.0, "bottom": 25.0, "left": 26.0, "right": 25.0}

    section.page_width = Mm(w_mm)
    section.page_height = Mm(h_mm)
    sect_pr = section._sectPr
    for old_grid in sect_pr.findall(qn("w:docGrid")):
        sect_pr.remove(old_grid)

    if zero_margins:
        section.top_margin = Mm(0)
        section.bottom_margin = Mm(0)
        section.left_margin = Mm(0)
        section.right_margin = Mm(0)
        section.header_distance = Mm(0)
        section.footer_distance = Mm(0)
        return int(round(w_mm * TWIPS_PER_MM)), int(round(h_mm * TWIPS_PER_MM))

    top = float(margins.get("top", 30.0))
    bottom = float(margins.get("bottom", 25.0))
    left = float(margins.get("left", 26.0))
    right = float(margins.get("right", 25.0))
    section.top_margin = Mm(top)
    section.bottom_margin = Mm(bottom)
    section.left_margin = Mm(left)
    section.right_margin = Mm(right)
    section.header_distance = Mm(15.0)
    section.footer_distance = Mm(13.0)
    pw = int(round((w_mm - left - right) * TWIPS_PER_MM))
    ph = int(round((h_mm - top - bottom) * TWIPS_PER_MM))

    grid_cfg = page_cfg.get("grid") or {}
    grid_lines = int(grid_cfg.get("lines", 22) or 22)
    if not landscape and grid_lines > 0:
        line_pitch = max(240, int(ph // grid_lines))
        sect_pr.append(parse_xml(f'<w:docGrid {nsdecls("w")} w:type="lines" w:linePitch="{line_pitch}"/>'))
    return pw, ph


def _clear_section_hf(section) -> None:
    section.different_first_page_header_footer = False
    for attr in (
        "header",
        "footer",
        "first_page_header",
        "first_page_footer",
        "even_page_header",
        "even_page_footer",
    ):
        part = getattr(section, attr, None)
        if part is None:
            continue
        part.is_linked_to_previous = False
        el = part._element
        for child in list(el):
            el.remove(child)
        el.append(OxmlElement("w:p"))


def _populate_footer_page_number(
    footer_part,
    align: WD_ALIGN_PARAGRAPH,
    *,
    prefix: str,
    suffix: str,
    font_name: str,
    size_pt: float,
    indent_side: str | None = None,
    indent_twips: int = 0,
    indent_chars: int = 0,
) -> None:
    footer_part.is_linked_to_previous = False
    el = footer_part._element
    for child in list(el):
        el.remove(child)

    p = footer_part.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Pt(0)
    if indent_side in {"left", "right"} and (indent_twips > 0 or indent_chars > 0):
        p_pr = p._p.get_or_add_pPr()
        for old_ind in p_pr.findall(qn("w:ind")):
            p_pr.remove(old_ind)
        p_pr.append(
            parse_xml(
                f'<w:ind {nsdecls("w")} w:firstLine="0" w:{indent_side}="{indent_twips}" w:{indent_side}Chars="{indent_chars}"/>'
            )
        )

    if prefix:
        r_pre = p.add_run(prefix)
        _set_run_font(r_pre, font_name, size_pt, bold=False)

    fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>')
    r_inner = fld.find(qn("w:r"))
    if r_inner is not None:
        _set_element_font(r_inner, font_name, size_pt, bold=False)
    p._p.append(fld)

    if suffix:
        r_suf = p.add_run(suffix)
        _set_run_font(r_suf, font_name, size_pt, bold=False)


def _apply_section_footer_page_number(
    section,
    fmt: dict[str, Any],
    *,
    restart_at: int | None = None,
) -> None:
    pn = fmt.get("page_numbers", {})
    prefix = str(pn.get("prefix", "-"))
    suffix = str(pn.get("suffix", "-"))
    font_name = str(pn.get("font", "宋体"))
    size_pt = float(pn.get("size_pt", 14.0))
    align = str(pn.get("align", "odd_right_even_left")).lower()
    odd_even = bool(pn.get("odd_even", align == "odd_right_even_left"))
    indent_chars_f = float(pn.get("indent_chars", 1.0))
    indent_twips = max(0, int(round(size_pt * 20 * indent_chars_f)))
    indent_chars_pct = max(0, int(round(indent_chars_f * 100)))

    sect_pr = section._sectPr
    for old in sect_pr.findall(qn("w:pgNumType")):
        sect_pr.remove(old)
    if restart_at is not None:
        sect_pr.append(parse_xml(f'<w:pgNumType {nsdecls("w")} w:start="{int(restart_at)}"/>'))

    section.different_first_page_header_footer = False
    if odd_even or align == "odd_right_even_left":
        doc_obj = getattr(section.part, "document", None)
        if doc_obj is not None:
            doc_obj.settings.odd_and_even_pages_header_footer = True
        # Odd pages: bottom-right indented by 1 character; Even pages: bottom-left indented by 1 character
        _populate_footer_page_number(
            section.footer,
            WD_ALIGN_PARAGRAPH.RIGHT,
            prefix=prefix,
            suffix=suffix,
            font_name=font_name,
            size_pt=size_pt,
            indent_side="right",
            indent_twips=indent_twips,
            indent_chars=indent_chars_pct,
        )
        _populate_footer_page_number(
            section.even_page_footer,
            WD_ALIGN_PARAGRAPH.LEFT,
            prefix=prefix,
            suffix=suffix,
            font_name=font_name,
            size_pt=size_pt,
            indent_side="left",
            indent_twips=indent_twips,
            indent_chars=indent_chars_pct,
        )
    else:
        target_align = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
        }.get(align, WD_ALIGN_PARAGRAPH.CENTER)
        _populate_footer_page_number(
            section.footer,
            target_align,
            prefix=prefix,
            suffix=suffix,
            font_name=font_name,
            size_pt=size_pt,
        )
        _populate_footer_page_number(
            section.even_page_footer,
            target_align,
            prefix=prefix,
            suffix=suffix,
            font_name=font_name,
            size_pt=size_pt,
        )


def _set_header_watermark(
    header_part,
    watermark_path: Path | None,
    width_mm: float,
    height_mm: float,
    doc_pr_counter: list[int],
) -> None:
    if header_part is None:
        return
    header_part.is_linked_to_previous = False
    el = header_part._element
    for child in list(el):
        el.remove(child)
    if not watermark_path or not Path(watermark_path).is_file():
        el.append(OxmlElement("w:p"))
        return

    with Image.open(watermark_path) as img:
        img_w, img_h = img.size
    if img_w <= 0 or img_h <= 0:
        el.append(OxmlElement("w:p"))
        return

    max_cx = int(round(width_mm * EMU_PER_MM))
    max_cy = int(round(height_mm * EMU_PER_MM))
    scale = min(max_cx / img_w, max_cy / img_h)
    cx = int(round(img_w * scale))
    cy = int(round(img_h * scale))
    off_x = max(0, (max_cx - cx) // 2)
    off_y = max(0, (max_cy - cy) // 2)

    r_id, _ = header_part.part.get_or_add_image(str(watermark_path))
    doc_pr_counter[0] += 1
    doc_id = doc_pr_counter[0]

    anchor_xml = f"""
    <w:r {nsdecls("w", "wp", "a", "pic", "r")}>
      <w:drawing>
        <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
                   relativeHeight="251657216" behindDoc="1" locked="0"
                   layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="page"><wp:posOffset>{off_x}</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="page"><wp:posOffset>{off_y}</wp:posOffset></wp:positionV>
          <wp:extent cx="{cx}" cy="{cy}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:wrapNone/>
          <wp:docPr id="{doc_id}" name="Watermark{doc_id}" descr="底图"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:nvPicPr>
                  <pic:cNvPr id="{doc_id}" name="Watermark{doc_id}" descr="底图"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="{r_id}"/>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing>
    </w:r>
    """
    p = header_part.add_paragraph()
    p._p.append(parse_xml(anchor_xml))


def _apply_section_watermark(
    section,
    watermark_path: Path | None,
    width_mm: float,
    height_mm: float,
    doc_pr_counter: list[int],
) -> None:
    if section.first_page_header is not None:
        section.first_page_header.is_linked_to_previous = False
        for ch in list(section.first_page_header._element):
            section.first_page_header._element.remove(ch)
        section.first_page_header._element.append(OxmlElement("w:p"))
    _set_header_watermark(section.header, watermark_path, width_mm, height_mm, doc_pr_counter)
    _set_header_watermark(section.even_page_header, watermark_path, width_mm, height_mm, doc_pr_counter)


def _insert_fullpage_image(
    doc: Document,
    img_path: Path,
    page_w_mm: float,
    page_h_mm: float,
    doc_pr_counter: list[int] | None = None,
) -> None:
    """Insert a full-bleed single-page image anchored to page (0,0) so it never overflows onto a 2nd page."""
    cx = int(round(page_w_mm * EMU_PER_MM))
    cy = int(round(page_h_mm * EMU_PER_MM))
    r_id, _ = doc.part.get_or_add_image(str(img_path))
    if doc_pr_counter is None:
        doc_pr_counter = [9000]
    doc_pr_counter[0] += 1
    doc_id = doc_pr_counter[0]
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    anchor_xml = f"""
    <w:r {nsdecls("w", "wp", "a", "pic", "r")}>
      <w:drawing>
        <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
                   relativeHeight="251658240" behindDoc="0" locked="0"
                   layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
          <wp:extent cx="{cx}" cy="{cy}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:wrapNone/>
          <wp:docPr id="{doc_id}" name="FullPage{doc_id}"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:nvPicPr>
                  <pic:cNvPr id="{doc_id}" name="FullPage{doc_id}"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="{r_id}"/>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing>
    </w:r>
    """
    p._p.append(parse_xml(anchor_xml))


def _add_scaled_image_paragraph(
    doc: Document,
    img_path: Path,
    printable_w_twips: int,
    printable_h_twips: int,
    *,
    caption: str = "",
) -> None:
    with Image.open(img_path) as im:
        iw, ih = im.size
    if iw <= 0 or ih <= 0:
        return
    max_w_mm = printable_w_twips / TWIPS_PER_MM * 0.96
    max_h_mm = printable_h_twips / TWIPS_PER_MM * 0.88
    if caption:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(6)
        cp.paragraph_format.space_after = Pt(4)
        cp.paragraph_format.keep_with_next = True
        cr = cp.add_run(caption)
        _set_run_font(cr, "方正黑体_GBK", 14.0, bold=True)

    ratio = iw / ih
    w_mm = max_w_mm
    h_mm = w_mm / ratio
    if h_mm > max_h_mm:
        h_mm = max_h_mm
        w_mm = h_mm * ratio

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Pt(0)
    p.add_run().add_picture(str(img_path), width=Mm(w_mm), height=Mm(h_mm))


_CN_STRUCT = r"[一二三四五六七八九十百零]+"
_NO_INDENT_ROLE_RE = re.compile(
    r"^(主\s*任|副主任|委\s*员|顾\s*问|秘\s*书\s*长|副秘书长|成\s*员|组\s*长|副组长|"
    r"比\s*赛\s*监\s*督|技\s*术\s*代\s*表|裁\s*判\s*长|副裁判长|裁\s*判\s*员|联\s*系\s*人|单位（盖章）|"
    r"参赛单位|参赛组别|领\s*队|教\s*练\s*员|电\s*话)\s*[：:]"
)


def _split_title_and_subtitle(title_text: str) -> tuple[str, str | None]:
    text = (title_text or "").strip()
    if "——" in text:
        parts = text.split("——", 1)
        if parts[0].strip() and parts[1].strip():
            return parts[0].strip(), f"——{parts[1].strip()}"
    return text, None


def _clean_title_punctuation(text: str) -> str:
    """Keep only allowed title punctuation (《》“”、-（）), stripping trailing period/colon."""
    return re.sub(r"[。：:;；.,，、]+$", "", (text or "").strip()).strip()


def _format_trapezoid_title_lines(text: str, max_chars: int = 18) -> list[str]:
    """Arrange overlong multi-line titles into trapezoid or diamond shape, never I-beam."""
    txt = (text or "").strip()
    n = len(txt)
    if n <= max_chars or n < 10:
        return [txt]

    def _split_score(idx: int, total: int, l1_target_diff: int) -> float:
        if idx < 4 or total - idx < 4:
            return -1000.0
        left_len, right_len = idx, total - idx
        if left_len > max_chars or right_len > max_chars:
            return -500.0
        if abs(left_len - right_len) <= 1:
            # Prefer trapezoidal mismatch over equal-width I-beam lines, without breaking words
            score = -15.0 if left_len == right_len else -6.0
        else:
            score = 10.0 - abs(abs(left_len - right_len) - l1_target_diff)
        prev_ch = txt[idx - 1]
        next_ch = txt[idx]
        # Never split inside ASCII runs, numbers, date/ordinal units, or brackets
        if (prev_ch.isascii() and prev_ch.isalnum() and next_ch.isascii() and next_ch.isalnum()):
            return -800.0
        if (prev_ch.isdigit() and next_ch in "年月日届期号第岁个名场次项") or (
            prev_ch == "第" and next_ch.isdigit()
        ):
            return -500.0
        if next_ch in "）)》」”】、，。":
            return -600.0
        if prev_ch in "（(《「“【":
            return -600.0
        if prev_ch in "、）)》” ":
            score += 25.0
        elif prev_ch in "会的于杯赛区市省局通知":
            score += 12.0
        elif txt[max(0, idx - 2):idx] in {"关于", "举行", "开展", "举办"}:
            score += 18.0
        return score

    if n <= max_chars * 2 - 2:
        # Prefer gentle trapezoid (longer top or longer bottom depending on natural boundary)
        best_idx = max(range(4, n - 3), key=lambda i: _split_score(i, n, 3))
        if _split_score(best_idx, n, 3) > -150.0:
            return [txt[:best_idx].strip(), txt[best_idx:].strip()]

    # Three-line diamond shape: short - longest - short
    target_mid = min(max_chars, max(10, int(round(n * 0.44))))
    side = (n - target_mid) // 2
    p1 = max(4, side)
    p2 = min(n - 4, p1 + target_mid)
    if p2 > p1 + 2 and (p2 - p1) > max(p1, n - p2):
        return [txt[:p1].strip(), txt[p1:p2].strip(), txt[p2:].strip()]
    return [txt]


def _make_blank_spacer_p_el(
    size_pt: float = 10.5,
    *,
    line_pt: float | None = None,
    keep_next: bool = False,
):
    p = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    line_val = int(round((line_pt if line_pt is not None else max(size_pt, 13.0)) * 20))
    p_pr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:before="0" w:after="0" w:line="{line_val}" w:lineRule="exact"/>'))
    p_pr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
    if keep_next:
        p_pr.append(OxmlElement("w:keepNext"))
    r_pr = OxmlElement("w:rPr")
    sz_val = str(int(round(size_pt * 2)))
    r_pr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz_val}"/>'))
    r_pr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{sz_val}"/>'))
    p_pr.append(r_pr)
    p.append(p_pr)
    r = OxmlElement("w:r")
    r_inner_pr = OxmlElement("w:rPr")
    r_inner_pr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz_val}"/>'))
    r_inner_pr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{sz_val}"/>'))
    r.append(r_inner_pr)
    r.append(OxmlElement("w:t"))
    p.append(r)
    return p


def _append_blank_spacer(
    doc: Document,
    size_pt: float = 10.5,
    *,
    line_pt: float | None = None,
    keep_next: bool = False,
    page_break_before: bool = False,
) -> None:
    el = _make_blank_spacer_p_el(size_pt, line_pt=line_pt, keep_next=keep_next)
    if page_break_before:
        p_pr = el.find(qn("w:pPr"))
        if p_pr is not None and p_pr.find(qn("w:pageBreakBefore")) is None:
            p_pr.append(OxmlElement("w:pageBreakBefore"))
    sect_pr = doc.element.body.find(qn("w:sectPr"))
    if sect_pr is not None:
        sect_pr.addprevious(el)
    else:
        doc.element.body.append(el)


def _rewrite_paragraph_text(p_el, new_text: str) -> None:
    """Safely replace plain text inside a paragraph while preserving bookmarks/rPr."""
    if (
        p_el.find(".//" + qn("w:drawing")) is not None
        or p_el.find(".//" + qn("w:pict")) is not None
        or p_el.find(".//" + qn("w:fldChar")) is not None
        or p_el.find(".//" + qn("w:hyperlink")) is not None
    ):
        return
    t_nodes = list(p_el.iter(qn("w:t")))
    if not t_nodes:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = new_text
        r.append(t)
        p_el.append(r)
        return
    t_nodes[0].text = new_text
    t_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in t_nodes[1:]:
        extra.text = ""


def _detect_and_normalize_structure(txt: str) -> tuple[str, str]:
    """Classify paragraph level per Official Format Rules 2-4 and fix erroneous punctuation."""
    s = (txt or "").strip()
    if not s:
        return s, "blank"

    # Rule 4: Attachment page marker ("附件" or "附件1", no colon, no inner space)
    m_att = re.match(r"^附\s*件(?:\s*(\d+))?\s*[：:]?\s*$", s)
    if m_att:
        num = m_att.group(1) or ""
        return f"附件{num}", "attachment_label"

    # Rule 3: Attachment explanation list ("附件：...")
    if re.match(r"^附\s*件\s*[：:]\s*\S+", s):
        return s, "attachment_desc"

    # Rule 2 Level 1: "一、"
    m_h1 = re.match(rf"^({_CN_STRUCT})\s*[、.．,，]\s*(\S.*)$", s)
    if m_h1:
        return f"{m_h1.group(1)}、{m_h1.group(2).strip()}", "heading1"

    # Rule 2 Level 2: "（一）" (strip erroneous trailing 、 or ， like "(一)、")
    m_h2 = re.match(rf"^[（(]\s*({_CN_STRUCT})\s*[）)]\s*[、，,．.]?\s*(\S.*)$", s)
    if m_h2:
        norm = f"（{m_h2.group(1)}）{m_h2.group(2).strip()}"
        if len(norm) <= 38 and "。" not in norm[:-1] and not re.search(r"[，；].*。", norm):
            return norm, "heading2"
        return norm, "body"

    # Rule 2 Level 3: "1．" (normalize erroneous 1、 / 1, / 1. -> 1． without touching decimals)
    m_h3 = re.match(r"^(\d{1,2})\s*[.．、，,]\s*([^\d\s.%％].*)$", s)
    if m_h3:
        return f"{m_h3.group(1)}．{m_h3.group(2).strip()}", "heading3"

    # Rule 2 Level 4: "（1）" (strip erroneous trailing 、 or ， like "(1)，")
    m_h4 = re.match(r"^[（(]\s*(\d{1,2})\s*[）)]\s*[、，,．.]?\s*(\S.*)$", s)
    if m_h4:
        return f"（{m_h4.group(1)}）{m_h4.group(2).strip()}", "heading4"

    return s, "body"


def _split_attachment_desc_items(txt: str) -> list[tuple[str, bool]]:
    """Parse '附件：...' into properly punctuated, trailing-punct-free items."""
    body = re.sub(r"^附\s*件\s*[：:]\s*", "", (txt or "").strip())
    body = re.sub(r"[。；;.,，、]+$", "", body).strip()
    numbered_parts = re.split(r"(?:^|\s+)(?=\d{1,2}\s*[.．、]\s*[^\d\s])", body)
    numbered_parts = [p.strip() for p in numbered_parts if p.strip()]
    if len(numbered_parts) >= 2 or re.match(r"^\d{1,2}\s*[.．、]\s*\S", body):
        out: list[tuple[str, bool]] = []
        for idx, part in enumerate(numbered_parts):
            clean = re.sub(r"^(\d{1,2})\s*[.．、，,]\s*", r"\1．", part)
            clean = re.sub(r"[。；;.,，、]+$", "", clean).strip()
            out.append((f"附件：{clean}" if idx == 0 else clean, idx == 0))
        return out
    return [(f"附件：{body}", True)]


def _set_paragraph_xml_layout(
    p_el,
    *,
    jc: str = "both",
    line_twips: int = 600,
    first_line_twips: int | None = 0,
    first_line_chars: int | None = None,
    left_twips: int | None = None,
    left_chars: int | None = None,
    hanging_twips: int | None = None,
    hanging_chars: int | None = None,
    keep_next: bool = False,
    page_break_before: bool | None = None,
) -> None:
    p_pr = p_el.find(qn("w:pPr"))
    if p_pr is None:
        p_pr = OxmlElement("w:pPr")
        p_el.insert(0, p_pr)
    for old_jc in p_pr.findall(qn("w:jc")):
        p_pr.remove(old_jc)
    p_pr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc}"/>'))

    for old_sp in p_pr.findall(qn("w:spacing")):
        p_pr.remove(old_sp)
    p_pr.append(
        parse_xml(
            f'<w:spacing {nsdecls("w")} w:before="0" w:after="0" w:line="{line_twips}" w:lineRule="exact"/>'
        )
    )

    if first_line_twips is not None or left_twips is not None or hanging_twips is not None:
        for old_ind in p_pr.findall(qn("w:ind")):
            p_pr.remove(old_ind)
        attrs = []
        if left_twips is not None:
            attrs.append(f'w:left="{left_twips}"')
        if left_chars is not None:
            attrs.append(f'w:leftChars="{left_chars}"')
        if hanging_twips is not None:
            attrs.append(f'w:hanging="{hanging_twips}"')
        elif first_line_twips is not None:
            attrs.append(f'w:firstLine="{first_line_twips}"')
        if hanging_chars is not None:
            attrs.append(f'w:hangingChars="{hanging_chars}"')
        elif first_line_chars is not None and hanging_twips is None:
            attrs.append(f'w:firstLineChars="{first_line_chars}"')
        p_pr.append(parse_xml(f'<w:ind {nsdecls("w")} {" ".join(attrs)}/>'))

    if keep_next:
        if p_pr.find(qn("w:keepNext")) is None:
            p_pr.append(OxmlElement("w:keepNext"))
    if page_break_before is True:
        if p_pr.find(qn("w:pageBreakBefore")) is None:
            p_pr.append(OxmlElement("w:pageBreakBefore"))
    elif page_break_before is False:
        for pb in p_pr.findall(qn("w:pageBreakBefore")):
            p_pr.remove(pb)


def _apply_role_style_to_p_el(
    p_el,
    role: str,
    fmt: dict[str, Any],
    *,
    keep_next: bool = False,
    page_break_before: bool | None = None,
    preserve_indent: bool = False,
    preserve_jc: str | None = None,
) -> None:
    fonts = fmt["fonts"]
    body_cfg = fonts["body"]
    body_pt = float(body_cfg.get("size_pt", 16.0))
    line_twips = int(round(float(body_cfg.get("line_spacing_pt", 30.0)) * 20))
    indent_chars = float(body_cfg.get("first_line_indent_chars", 2.0))
    indent_twips = int(round(body_pt * 20 * indent_chars))
    indent_chars_pct = int(round(indent_chars * 100))

    if role in {"title", "attachment_title"}:
        cfg = fonts["title"]
        fn, sz, bold = str(cfg.get("name", "方正小标宋_GBK")), float(cfg.get("size_pt", 22.0)), bool(cfg.get("bold", False))
        jc, fl_tw, fl_ch = "center", 0, 0
        left_tw = left_ch = hang_tw = hang_ch = None
        keep_next = True
    elif role == "subtitle":
        cfg = fonts.get("subtitle", {})
        fn, sz, bold = str(cfg.get("name", "方正楷体_GBK")), float(cfg.get("size_pt", 16.0)), bool(cfg.get("bold", False))
        jc, fl_tw, fl_ch = "center", 0, 0
        left_tw = left_ch = hang_tw = hang_ch = None
        keep_next = True
    elif role == "attachment_label":
        cfg = fonts.get("attachment_label", fonts["heading1"])
        fn, sz, bold = str(cfg.get("name", "方正黑体_GBK")), float(cfg.get("size_pt", 16.0)), bool(cfg.get("bold", False))
        jc, fl_tw, fl_ch = "left", 0, 0
        left_tw = left_ch = hang_tw = hang_ch = None
        keep_next = True
    elif role == "heading1":
        cfg = fonts["heading1"]
        fn, sz, bold = str(cfg.get("name", "方正黑体_GBK")), float(cfg.get("size_pt", 16.0)), bool(cfg.get("bold", False))
        jc, fl_tw, fl_ch = "both", indent_twips, indent_chars_pct
        left_tw = left_ch = hang_tw = hang_ch = None
        keep_next = True
    elif role == "heading2":
        cfg = fonts["heading2"]
        fn, sz, bold = str(cfg.get("name", "方正楷体_GBK")), float(cfg.get("size_pt", 16.0)), bool(cfg.get("bold", False))
        jc, fl_tw, fl_ch = "both", indent_twips, indent_chars_pct
        left_tw = left_ch = hang_tw = hang_ch = None
    elif role in {"heading3", "heading4"}:
        cfg = fonts.get(role, body_cfg)
        fn, sz, bold = str(cfg.get("name", "方正仿宋_GBK")), float(cfg.get("size_pt", 16.0)), bool(cfg.get("bold", False))
        jc, fl_tw, fl_ch = "both", indent_twips, indent_chars_pct
        left_tw = left_ch = hang_tw = hang_ch = None
    elif role == "attachment_desc_single":
        fn, sz, bold = str(body_cfg.get("name", "方正仿宋_GBK")), body_pt, False
        char_tw = int(round(body_pt * 20))
        jc, fl_tw, fl_ch = "both", None, None
        # Twips-only hanging indent: avoid LibreOffice compounding w:leftChars on top of w:left
        left_tw, left_ch = char_tw * 5, None
        hang_tw, hang_ch = char_tw * 3, None
    elif role == "attachment_desc_multi_first":
        fn, sz, bold = str(body_cfg.get("name", "方正仿宋_GBK")), body_pt, False
        char_tw = int(round(body_pt * 20))
        # Left start of "附件：": 2 chars; "附件："=3 chars; "1．"=1.5 chars -> wrap under first char at 6.5 chars
        jc, fl_tw, fl_ch = "both", None, None
        left_tw, left_ch = int(round(char_tw * 6.5)), None
        hang_tw, hang_ch = int(round(char_tw * 4.5)), None
    elif role == "attachment_desc_multi_next":
        fn, sz, bold = str(body_cfg.get("name", "方正仿宋_GBK")), body_pt, False
        char_tw = int(round(body_pt * 20))
        # Align "2．" with "1．" at 5 chars, wrapped text at 6.5 chars
        jc, fl_tw, fl_ch = "both", None, None
        left_tw, left_ch = int(round(char_tw * 6.5)), None
        hang_tw, hang_ch = int(round(char_tw * 1.5)), None
    else:
        fn, sz, bold = str(body_cfg.get("name", "方正仿宋_GBK")), body_pt, bool(body_cfg.get("bold", False))
        jc = preserve_jc if preserve_jc in {"center", "right"} else "both"
        if preserve_indent or preserve_jc in {"center", "right"}:
            fl_tw, fl_ch = (0, 0) if preserve_jc in {"center", "right"} else (None, None)
        else:
            fl_tw, fl_ch = indent_twips, indent_chars_pct
        left_tw = left_ch = hang_tw = hang_ch = None

    for r in p_el.findall(qn("w:r")):
        _set_element_font(r, fn, sz, bold=bold)
    p_pr = p_el.find(qn("w:pPr"))
    if p_pr is not None and p_pr.find(qn("w:rPr")) is not None:
        _set_element_font(p_pr, fn, sz, bold=bold)

    _set_paragraph_xml_layout(
        p_el,
        jc=jc,
        line_twips=line_twips,
        first_line_twips=fl_tw,
        first_line_chars=fl_ch,
        left_twips=left_tw,
        left_chars=left_ch,
        hanging_twips=hang_tw,
        hanging_chars=hang_ch,
        keep_next=keep_next,
        page_break_before=page_break_before,
    )


def _add_bookmarked_chapter_heading(
    doc: Document,
    title: str,
    anchor: str,
    bookmark_id: int,
    fmt: dict[str, Any],
    *,
    centered: bool = True,
    page_break_before: bool = False,
) -> None:
    title_cfg = fmt["fonts"]["title"]
    top_blank = float(title_cfg.get("top_blank_size_pt", 10.5))
    bot_blank = float(title_cfg.get("bottom_blank_size_pt", 10.5))
    broke = False
    if centered and top_blank > 0:
        _append_blank_spacer(doc, top_blank, line_pt=13.0, keep_next=True, page_break_before=page_break_before)
        broke = True

    main_title, subtitle = _split_title_and_subtitle(title)
    if centered:
        main_title = _clean_title_punctuation(main_title)
        if subtitle:
            subtitle = f"——{_clean_title_punctuation(subtitle[2:])}"
    p = doc.add_paragraph()
    hcfg = title_cfg if centered else fmt["fonts"]["heading1"]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    line_pt = float(fmt["fonts"]["body"].get("line_spacing_pt", 30.0))
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(line_pt)
    pf.first_line_indent = Pt(0)
    pf.keep_with_next = True
    if page_break_before and not broke:
        pf.page_break_before = True

    p_pr = p._p.get_or_add_pPr()
    for old_ol in p_pr.findall(qn("w:outlineLvl")):
        p_pr.remove(old_ol)
    p_pr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))

    bm_start = OxmlElement("w:bookmarkStart")
    bm_start.set(qn("w:id"), str(bookmark_id))
    bm_start.set(qn("w:name"), anchor)
    bm_end = OxmlElement("w:bookmarkEnd")
    bm_end.set(qn("w:id"), str(bookmark_id))

    lines = _format_trapezoid_title_lines(main_title) if centered else [main_title]
    for li, ln_txt in enumerate(lines):
        if li > 0:
            p.add_run().add_break()
        run = p.add_run(ln_txt)
        _set_run_font(
            run,
            str(hcfg.get("name", "方正小标宋_GBK" if centered else "方正黑体_GBK")),
            float(hcfg.get("size_pt", 22.0 if centered else 16.0)),
            bold=bool(hcfg.get("bold", False)),
        )
    p._p.insert(0, bm_start)
    p._p.append(bm_end)

    if subtitle and centered:
        sub_cfg = fmt["fonts"].get("subtitle", {})
        sp = doc.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        spf = sp.paragraph_format
        spf.space_before = Pt(0)
        spf.space_after = Pt(0)
        spf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        spf.line_spacing = Pt(line_pt)
        spf.first_line_indent = Pt(0)
        spf.keep_with_next = True
        sr = sp.add_run(subtitle)
        _set_run_font(
            sr,
            str(sub_cfg.get("name", "方正楷体_GBK")),
            float(sub_cfg.get("size_pt", 16.0)),
            bold=bool(sub_cfg.get("bold", False)),
        )

    if centered and bot_blank > 0:
        _append_blank_spacer(doc, bot_blank, line_pt=13.0, keep_next=True)


def _add_toc_block(
    doc: Document,
    toc_entries: list[dict[str, Any]],
    printable_w_twips: int,
    fmt: dict[str, Any],
) -> None:
    toc_cfg = fmt["fonts"].get("toc", {})
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(6)
    title_p.paragraph_format.space_after = Pt(18)
    tr = title_p.add_run(str(toc_cfg.get("title", "目    录")))
    _set_run_font(
        tr,
        str(toc_cfg.get("title_name", "方正小标宋_GBK")),
        float(toc_cfg.get("title_size_pt", 22.0)),
        bold=False,
    )

    entry_font = str(toc_cfg.get("entry_name", "方正仿宋_GBK"))
    entry_size = float(toc_cfg.get("entry_size_pt", 14.0))
    line_pt = float(toc_cfg.get("line_spacing_pt", 24.0))

    for item in toc_entries:
        label = str(item.get("toc_label") or item.get("title") or "")
        anchor = str(item.get("anchor") or "")
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(2)
        pf.space_after = Pt(4)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_pt)
        pf.first_line_indent = Pt(0)

        p_pr = p._p.get_or_add_pPr()
        tabs = parse_xml(
            f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="{printable_w_twips}"/></w:tabs>'
        )
        p_pr.append(tabs)

        r_pr_xml = (
            f'<w:rPr {nsdecls("w")}>'
            f'<w:rFonts w:ascii="{entry_font}" w:eastAsia="{entry_font}" w:hAnsi="{entry_font}" w:cs="{entry_font}"/>'
            f'<w:sz w:val="{int(round(entry_size * 2))}"/><w:szCs w:val="{int(round(entry_size * 2))}"/>'
            f"</w:rPr>"
        )
        hyperlink = parse_xml(
            f'<w:hyperlink {nsdecls("w")} w:anchor="{anchor}" w:history="1">'
            f"<w:r>{r_pr_xml}<w:t>{_xml_escape(label)}</w:t></w:r>"
            f"<w:r>{r_pr_xml}<w:tab/></w:r>"
            f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> PAGEREF {anchor} \\h </w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f"<w:r>{r_pr_xml}<w:t>1</w:t></w:r>"
            f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            f"</w:hyperlink>"
        )
        p._p.append(hyperlink)


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _table_border_xml() -> str:
    return (
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="333333"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="333333"/>'
        "</w:tblBorders>"
    )


def _excel_serial_to_cn_date(val: float | int) -> str | None:
    try:
        iv = int(round(float(val)))
        if not (36526 <= iv <= 65000):
            return None
        d = dt.date(1899, 12, 30) + dt.timedelta(days=iv)
        return f"{d.month}月{d.day}日"
    except Exception:
        return None


def _is_date_number_format(num_fmt: str | None) -> bool:
    if not num_fmt:
        return False
    low = num_fmt.lower()
    if low == "general" or "@" in low:
        return False
    return any(ch in low for ch in ("yy", "mm", "dd", "月", "日"))


def _format_cell_value(cell, *, is_date_col: bool = False) -> str:
    val = cell.value
    if val is None:
        return ""
    if isinstance(val, dt.datetime):
        if val.hour == 0 and val.minute == 0 and val.second == 0:
            return f"{val.month}月{val.day}日"
        return f"{val.month}月{val.day}日 {val.hour}:{val.minute:02d}"
    if isinstance(val, dt.date):
        return f"{val.month}月{val.day}日"
    if isinstance(val, dt.time):
        if val.second:
            return f"{val.hour}:{val.minute:02d}:{val.second:02d}"
        return f"{val.hour}:{val.minute:02d}"
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if (is_date_col or _is_date_number_format(getattr(cell, "number_format", None))) and 36526 <= val <= 65000:
            converted = _excel_serial_to_cn_date(val)
            if converted:
                return converted
        if float(val).is_integer():
            return str(int(round(float(val))))
        return f"{val:g}"
    return str(val).strip()


def _detect_date_columns(ws) -> set[int]:
    date_cols: set[int] = set()
    max_scan = min(ws.max_row or 0, 8)
    for r in range(1, max_scan + 1):
        for c in range(1, (ws.max_column or 0) + 1):
            v = str(ws.cell(r, c).value or "").strip()
            if v in {"日期", "比赛日期", "时间/日期", "比赛时间"} and c == 1:
                date_cols.add(c)
            elif "日期" in v:
                date_cols.add(c)
    return date_cols


from contextlib import contextmanager


@contextmanager
def _open_xlsx_workbook(src_path: Path, *, read_only: bool = False):
    src_path = Path(src_path)
    if src_path.suffix.lower() == ".xlsx":
        wb = openpyxl.load_workbook(src_path, data_only=True, read_only=read_only)
        try:
            yield wb
        finally:
            wb.close()
        return
    with tempfile.TemporaryDirectory(prefix="xls-conv-") as tmp_dir:
        converted = _run_soffice_convert(src_path, Path(tmp_dir), "xlsx")
        wb = openpyxl.load_workbook(converted, data_only=True, read_only=read_only)
        try:
            yield wb
        finally:
            wb.close()


def xlsx_sheet_dimensions(xlsx_path: Path, sheet_names: list[str] | None = None) -> int:
    """Return max active non-empty column count across selected sheets (used for landscape='auto')."""
    max_cols = 0
    with _open_xlsx_workbook(Path(xlsx_path), read_only=True) as wb:
        names = sheet_names or list(wb.sheetnames)
        for name in names:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            for row in ws.iter_rows(values_only=True):
                non_empty = [i + 1 for i, v in enumerate(row) if v not in (None, "")]
                if non_empty:
                    max_cols = max(max_cols, max(non_empty))
    return max_cols


def _emit_xlsx_subtable(
    doc: Document,
    ws,
    min_r: int,
    max_r: int,
    max_c: int,
    top_left_map: dict[tuple[int, int], tuple[int, int]],
    skip_merge_cells: set[tuple[int, int]],
    date_cols: set[int],
    printable_w_twips: int,
    fmt: dict[str, Any],
) -> None:
    """Render one contiguous non-empty Excel block as a styled fixed-grid Word table."""
    tbl_cfg = fmt["fonts"].get("table", {})
    body_font = str(tbl_cfg.get("name", "方正仿宋_GBK"))
    hdr_font = str(tbl_cfg.get("header_name", "方正黑体_GBK"))
    font_size = float(tbl_cfg.get("size_pt", 10.5))
    hdr_size = float(tbl_cfg.get("header_size_pt", font_size))

    # Trim trailing empty columns specific to this block, plus any leading empty/banner-spanned rows
    eff_max_c = 0
    for r in range(min_r, max_r + 1):
        for c in range(1, max_c + 1):
            if ws.cell(r, c).value not in (None, ""):
                eff_max_c = max(eff_max_c, c)
    if eff_max_c == 0:
        return

    while min_r <= max_r and all(
        (min_r, c) in skip_merge_cells
        or not _format_cell_value(ws.cell(min_r, c), is_date_col=c in date_cols).strip()
        for c in range(1, eff_max_c + 1)
    ):
        min_r += 1

    n_rows = max_r - min_r + 1
    if n_rows <= 0:
        return

    # Compute content-proportional column widths from 90th-percentile + capped max visual length
    col_scores = [3.5] * eff_max_c
    for cc in range(1, eff_max_c + 1):
        lengths: list[float] = []
        for rr in range(min_r, max_r + 1):
            if (rr, cc) in top_left_map and top_left_map[(rr, cc)][1] > cc:
                continue
            txt = _format_cell_value(ws.cell(rr, cc), is_date_col=cc in date_cols)
            if txt:
                # Longest line inside multi-line cell
                line_len = max(
                    sum(1.8 if ord(ch) > 127 else 1.0 for ch in ln)
                    for ln in (txt.splitlines() or [txt])
                )
                lengths.append(line_len)
        if lengths:
            lengths.sort()
            p90 = lengths[min(len(lengths) - 1, int(len(lengths) * 0.9))]
            col_scores[cc - 1] = max(3.5, min(p90 + 2.2, 34.0))
    score_sum = sum(col_scores) or float(eff_max_c)
    col_widths = [max(480, int(round(printable_w_twips * (sc / score_sum)))) for sc in col_scores]
    col_widths[-1] += printable_w_twips - sum(col_widths)

    # Determine header rows (typically row 0, or row 0..1 if row 0 contains vertical merge)
    header_row_end = min_r
    if any(top_left_map.get((min_r, cc), (min_r, cc))[0] > min_r for cc in range(1, eff_max_c + 1)):
        header_row_end = min(max_r, min_r + 1)

    table = doc.add_table(rows=n_rows, cols=eff_max_c)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    for old_b in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(old_b)
    tbl_pr.append(parse_xml(_table_border_xml()))
    for old_w in tbl_pr.findall(qn("w:tblW")):
        tbl_pr.remove(old_w)
    tbl_pr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{printable_w_twips}" w:type="dxa"/>'))
    for old_lay in tbl_pr.findall(qn("w:tblLayout")):
        tbl_pr.remove(old_lay)
    tbl_pr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

    # Rebuild tblGrid explicitly so LibreOffice & Word honor proportional col_widths strictly
    tbl_grid = table._tbl.find(qn("w:tblGrid"))
    if tbl_grid is not None:
        table._tbl.remove(tbl_grid)
    new_grid = OxmlElement("w:tblGrid")
    for w_tw in col_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w_tw))
        new_grid.append(gc)
    tbl_pr.addnext(new_grid)

    eff_size = font_size
    if eff_max_c >= 9:
        eff_size = min(eff_size, 9.5)
    elif eff_max_c >= 7:
        eff_size = min(eff_size, 10.0)

    for r_offset in range(n_rows):
        excel_r = min_r + r_offset
        tr = table.rows[r_offset]._tr
        tr_pr = tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        is_header = excel_r <= header_row_end
        if is_header:
            tr_pr.append(OxmlElement("w:tblHeader"))
        for c_idx in range(eff_max_c):
            excel_c = c_idx + 1
            cell = table.cell(r_offset, c_idx)
            cell.width = Emu(col_widths[c_idx] * EMU_PER_TWIP)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tc_pr = cell._tc.get_or_add_tcPr()
            for old_tcw in tc_pr.findall(qn("w:tcW")):
                tc_pr.remove(old_tcw)
            tc_pr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{col_widths[c_idx]}" w:type="dxa"/>'))
            if is_header:
                tc_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F2F4F7"/>'))
            if (excel_r, excel_c) in skip_merge_cells:
                continue
            text = _format_cell_value(ws.cell(excel_r, excel_c), is_date_col=excel_c in date_cols)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.first_line_indent = Pt(0)
            run = p.add_run(text)
            _set_run_font(
                run,
                hdr_font if is_header else body_font,
                hdr_size if is_header and eff_size >= 10 else eff_size,
                bold=is_header,
            )

    for (r0, c0), (r1, c1) in top_left_map.items():
        if r0 < min_r or r0 > max_r or c0 > eff_max_c:
            continue
        row0_idx = r0 - min_r
        row1_idx = min(r1, max_r) - min_r
        c0_idx = c0 - 1
        c1_idx = min(c1, eff_max_c) - 1
        if row1_idx > row0_idx or c1_idx > c0_idx:
            try:
                table.cell(row0_idx, c0_idx).merge(table.cell(row1_idx, c1_idx))
            except Exception:
                pass


def _import_xlsx_chapter(
    doc: Document,
    src_path: Path,
    printable_w_twips: int,
    fmt: dict[str, Any],
    *,
    sheet_names: list[str] | None = None,
    chapter_title: str = "",
) -> None:
    """Convert selected Excel worksheets into clean Word tables, skipping blank rows & repeating headers."""
    tbl_cfg = fmt["fonts"].get("table", {})
    hdr_font = str(tbl_cfg.get("header_name", "方正黑体_GBK"))
    clean_chap_title = re.sub(r"^[一二三四五六七八九十百]+[、.．]", "", re.sub(r"\s+", "", chapter_title))

    with _open_xlsx_workbook(src_path, read_only=False) as wb:
        names = [n for n in (sheet_names or wb.sheetnames) if n in wb.sheetnames]
        for s_idx, sname in enumerate(names):
            ws = wb[sname]
            max_r, max_c = 0, 0
            for r in range(1, (ws.max_row or 0) + 1):
                for c in range(1, (ws.max_column or 0) + 1):
                    if ws.cell(r, c).value not in (None, ""):
                        max_r = max(max_r, r)
                        max_c = max(max_c, c)
            if max_r == 0 or max_c == 0:
                continue

            merges = [
                rng
                for rng in ws.merged_cells.ranges
                if rng.max_row >= 1 and rng.min_row <= max_r and rng.min_col <= max_c
            ]
            top_left_map: dict[tuple[int, int], tuple[int, int]] = {}
            skip_merge_cells: set[tuple[int, int]] = set()
            row_spanned: set[int] = set()
            for rng in merges:
                r0 = max(rng.min_row, 1)
                r1 = min(rng.max_row, max_r)
                c0 = rng.min_col
                c1 = min(rng.max_col, max_c)
                if r1 > r0 or c1 > c0:
                    top_left_map[(r0, c0)] = (r1, c1)
                    for rr in range(r0, r1 + 1):
                        if rr > r0:
                            row_spanned.add(rr)
                        for cc in range(c0, c1 + 1):
                            if (rr, cc) != (r0, c0):
                                skip_merge_cells.add((rr, cc))

            date_cols = _detect_date_columns(ws)

            def _row_non_empty(rr: int) -> list[str]:
                return [
                    _format_cell_value(ws.cell(rr, c), is_date_col=c in date_cols).strip()
                    for c in range(1, max_c + 1)
                    if _format_cell_value(ws.cell(rr, c), is_date_col=c in date_cols).strip()
                ]

            def _is_banner_row(rr: int) -> str | None:
                vals = _row_non_empty(rr)
                if len(vals) != 1:
                    return None
                span = top_left_map.get((rr, 1))
                # Wide horizontal merge or single non-merged label row spanning empty siblings
                if (span and span[0] == rr and span[1] >= 2) or (max_c >= 3 and len(vals[0]) >= 2):
                    return vals[0]
                return None

            # Find contiguous non-empty row ranges separated by empty separator rows
            blocks: list[tuple[int, int]] = []
            cur_start: int | None = None
            for rr in range(1, max_r + 1):
                has_data = bool(_row_non_empty(rr)) or (rr in row_spanned)
                if has_data:
                    if cur_start is None:
                        cur_start = rr
                else:
                    if cur_start is not None:
                        blocks.append((cur_start, rr - 1))
                        cur_start = None
            if cur_start is not None:
                blocks.append((cur_start, max_r))

            if s_idx > 0:
                bp = doc.add_paragraph()
                bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                bp.paragraph_format.space_before = Pt(8)
                bp.paragraph_format.space_after = Pt(4)
                bp.paragraph_format.first_line_indent = Pt(0)
                bp.paragraph_format.keep_with_next = True
                _set_run_font(bp.add_run(sname), hdr_font, 14.0, bold=True)

            for b_start, b_end in blocks:
                # Extract leading banner/sub-title rows (including multi-row banner merges like A1:H2)
                banners: list[str] = []
                while b_start < b_end:
                    btxt = _is_banner_row(b_start)
                    if not btxt:
                        break
                    clean_b = re.sub(r"\s+", "", btxt)
                    if not (clean_chap_title and clean_b == clean_chap_title):
                        banners.append(btxt)
                    span = top_left_map.get((b_start, 1))
                    b_start = max(b_start + 1, (span[0] + 1) if span else (b_start + 1))
                    while b_start < b_end and not _row_non_empty(b_start):
                        b_start += 1

                for idx_b, btxt in enumerate(banners):
                    bp = doc.add_paragraph()
                    bp.alignment = WD_ALIGN_PARAGRAPH.CENTER if (len(btxt) >= 6 or b_start <= 4) else WD_ALIGN_PARAGRAPH.LEFT
                    bp.paragraph_format.space_before = Pt(8 if idx_b == 0 else 2)
                    bp.paragraph_format.space_after = Pt(4)
                    bp.paragraph_format.first_line_indent = Pt(0)
                    bp.paragraph_format.keep_with_next = True
                    _set_run_font(bp.add_run(btxt), hdr_font, 14.0 if idx_b == 0 else 12.0, bold=True)

                _emit_xlsx_subtable(
                    doc,
                    ws,
                    b_start,
                    b_end,
                    max_c,
                    top_left_map,
                    skip_merge_cells,
                    date_cols,
                    printable_w_twips,
                    fmt,
                )
                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_before = Pt(0)
                spacer.paragraph_format.space_after = Pt(4)


def _remap_and_sanitize_element(
    el,
    src_part,
    dst_doc: Document,
    doc_pr_counter: list[int],
    printable_w_twips: int,
    printable_h_twips: int,
) -> Any:
    """Deep-copy an OOXML paragraph or table, rebinding embedded image rels cleanly to dst_doc."""
    cloned = copy.deepcopy(el)

    # Strip internal sectPr inside imported paragraphs
    for sect_pr in cloned.findall(".//" + qn("w:sectPr")):
        parent = sect_pr.getparent()
        if parent is not None:
            parent.remove(sect_pr)

    # Strip conflicting bookmarks and proofErr noise
    for tag in ("w:bookmarkStart", "w:bookmarkEnd", "w:proofErr"):
        for node in cloned.findall(".//" + qn(tag)):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)

    # Remap DrawingML images (<a:blip r:embed="...">)
    rid_cache: dict[str, str] = {}
    for blip in cloned.iter(qn("a:blip")):
        old_rid = blip.get(qn("r:embed"))
        if not old_rid or old_rid not in src_part.rels:
            continue
        if old_rid not in rid_cache:
            blob = src_part.rels[old_rid].target_part.blob
            new_rid, _ = dst_doc.part.get_or_add_image(io.BytesIO(blob))
            rid_cache[old_rid] = new_rid
        blip.set(qn("r:embed"), rid_cache[old_rid])

    # Remap VML images (<v:imagedata r:id="...">) if any
    V_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"
    for v_img in cloned.iter(V_IMAGEDATA):
        old_rid = v_img.get(qn("r:id"))
        if not old_rid or old_rid not in src_part.rels:
            continue
        if old_rid not in rid_cache:
            blob = src_part.rels[old_rid].target_part.blob
            new_rid, _ = dst_doc.part.get_or_add_image(io.BytesIO(blob))
            rid_cache[old_rid] = new_rid
        v_img.set(qn("r:id"), rid_cache[old_rid])

    # Assign unique docPr ids & clamp image dimensions to fit within printable margins
    max_cx = int(round(printable_w_twips * EMU_PER_TWIP * 0.96))
    max_cy = int(round(printable_h_twips * EMU_PER_TWIP * 0.88))
    WP_DOCPR = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"
    WP_EXTENT = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
    for dp in cloned.iter(WP_DOCPR):
        doc_pr_counter[0] += 1
        dp.set("id", str(doc_pr_counter[0]))
    for ext in cloned.iter(WP_EXTENT):
        try:
            cx = int(ext.get("cx", "0"))
            cy = int(ext.get("cy", "0"))
        except ValueError:
            continue
        if cx > 0 and cy > 0 and (cx > max_cx or cy > max_cy):
            scale = min(max_cx / cx, max_cy / cy)
            new_cx = int(round(cx * scale))
            new_cy = int(round(cy * scale))
            ext.set("cx", str(new_cx))
            ext.set("cy", str(new_cy))
            drawing = ext.getparent()
            if drawing is not None:
                for a_ext in drawing.iter(qn("a:ext")):
                    a_ext.set("cx", str(new_cx))
                    a_ext.set("cy", str(new_cy))

    return cloned


def _has_media_el(p_el) -> bool:
    return (
        p_el.find(".//" + qn("a:blip")) is not None
        or p_el.find(".//{urn:schemas-microsoft-com:vml}imagedata") is not None
    )


def _xml_jc(p_el) -> str:
    p_pr = p_el.find(qn("w:pPr"))
    if p_pr is not None:
        jc_el = p_pr.find(qn("w:jc"))
        if jc_el is not None:
            val = (jc_el.get(qn("w:val")) or "").lower()
            if val == "start":
                return "left"
            if val == "end":
                return "right"
            return val
    return ""


def _xml_ind(p_el) -> tuple[int, int]:
    p_pr = p_el.find(qn("w:pPr"))
    if p_pr is None:
        return 0, 0
    ind = p_pr.find(qn("w:ind"))
    if ind is None:
        return 0, 0
    fl = ind.get(qn("w:firstLine")) or "0"
    lf = ind.get(qn("w:left")) or ind.get(qn("w:start")) or "0"
    try:
        fl_i = int(fl)
    except ValueError:
        fl_i = 0
    try:
        lf_i = int(lf)
    except ValueError:
        lf_i = 0
    return fl_i, lf_i


def _should_preserve_body_indent(txt: str, fl_twips: int, left_twips: int, *, chapter_kind: str = "") -> bool:
    if fl_twips >= 800 or left_twips >= 640:
        return True
    if re.search(r"年\s*月\s*日$", txt) or txt in {"（单位公章）", "(单位公章)"}:
        return True
    if _NO_INDENT_ROLE_RE.match(txt) or "□" in txt:
        return True
    if re.search(r"[：:]\s{3,}\S", txt):
        return True
    if chapter_kind == "organization" and len(txt) <= 44 and re.search(r"\S\s{2,}\S", txt):
        return True
    return False


def _style_imported_paragraph(
    p_el,
    fmt: dict[str, Any],
    *,
    is_roster: bool = False,
    role_override: str | None = None,
    keep_next: bool = False,
    page_break_before: bool | None = None,
    chapter_kind: str = "",
) -> None:
    """Apply official-document fonts, hierarchy, punctuation, and spacing cleanly."""
    if is_roster:
        return  # Roster paragraphs were already formatted by roster_skill
    has_img = _has_media_el(p_el)
    txt = "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()
    if has_img and not txt:
        p_pr = p_el.find(qn("w:pPr"))
        if p_pr is None:
            p_pr = OxmlElement("w:pPr")
            p_el.insert(0, p_pr)
        for old_jc in p_pr.findall(qn("w:jc")):
            p_pr.remove(old_jc)
        p_pr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
        for old_ind in p_pr.findall(qn("w:ind")):
            p_pr.remove(old_ind)
        return

    if not txt:
        line_twips = int(round(float(fmt["fonts"]["body"].get("line_spacing_pt", 30.0)) * 20))
        if role_override == "blank_10_5":
            _set_paragraph_xml_layout(p_el, jc="both", line_twips=260, first_line_twips=0, first_line_chars=0, keep_next=keep_next, page_break_before=False)
            for r in p_el.findall(qn("w:r")):
                _set_element_font(r, str(fmt["fonts"]["body"].get("name", "方正仿宋_GBK")), 10.5, bold=False)
        elif role_override == "blank_line":
            _set_paragraph_xml_layout(p_el, jc="both", line_twips=line_twips, first_line_twips=0, first_line_chars=0, keep_next=keep_next, page_break_before=False)
            for r in p_el.findall(qn("w:r")):
                _set_element_font(r, str(fmt["fonts"]["body"].get("name", "方正仿宋_GBK")), float(fmt["fonts"]["body"].get("size_pt", 16.0)), bold=False)
        return

    orig_jc = _xml_jc(p_el)
    orig_fl, orig_left = _xml_ind(p_el)

    if role_override:
        role = role_override
        norm_txt = txt
    else:
        norm_txt, role = _detect_and_normalize_structure(txt)
        if norm_txt != txt:
            _rewrite_paragraph_text(p_el, norm_txt)

    if role == "heading2" and len(norm_txt) <= 28:
        keep_next = True

    preserve_jc = orig_jc if (role == "body" and orig_jc in {"center", "right"} and len(norm_txt) < 36) else None
    preserve_indent = (role == "body") and _should_preserve_body_indent(
        norm_txt, orig_fl, orig_left, chapter_kind=chapter_kind
    )
    if role == "body" and preserve_jc == "center" and re.match(r"^——\s*\S+", norm_txt):
        role = "subtitle"

    _apply_role_style_to_p_el(
        p_el,
        role,
        fmt,
        keep_next=keep_next,
        page_break_before=page_break_before,
        preserve_indent=preserve_indent,
        preserve_jc=preserve_jc,
    )


def _style_imported_table(tbl_el, printable_w_twips: int, fmt: dict[str, Any], *, is_roster: bool = False) -> None:
    if is_roster:
        return
    tbl_pr = tbl_el.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl_el.insert(0, tbl_pr)
    for old_w in tbl_pr.findall(qn("w:tblW")):
        tbl_pr.remove(old_w)
    tbl_pr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{printable_w_twips}" w:type="dxa"/>'))
    if tbl_pr.find(qn("w:tblBorders")) is None:
        tbl_pr.append(parse_xml(_table_border_xml()))

    # Proportionally rescale tblGrid if total exceeds or widely differs from printable width
    grid = tbl_el.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        widths = [int(c.get(qn("w:w"), "0") or 0) for c in cols]
        total = sum(widths)
        if cols and total > 0 and total != printable_w_twips:
            scale = printable_w_twips / total
            new_ws = [max(300, int(round(w * scale))) for w in widths]
            new_ws[-1] += printable_w_twips - sum(new_ws)
            for col_el, nw in zip(cols, new_ws):
                col_el.set(qn("w:w"), str(nw))

    tbl_font = str(fmt["fonts"].get("table", {}).get("name", "方正仿宋_GBK"))
    for r in tbl_el.iter(qn("w:r")):
        _set_element_font(r, tbl_font, None)

    trs = tbl_el.findall(qn("w:tr"))
    for tr in trs:
        tr_pr = tr.find(qn("w:trPr"))
        if tr_pr is None:
            tr_pr = OxmlElement("w:trPr")
            tr.insert(0, tr_pr)
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))

    # Keep compact matrix tables (e.g., 6-7 row standings tables with VML diagonal lines) on one page
    if len(trs) <= 14:
        for tr in trs[:-1]:
            for p_el in tr.iter(qn("w:p")):
                p_pr = p_el.find(qn("w:pPr"))
                if p_pr is None:
                    p_pr = OxmlElement("w:pPr")
                    p_el.insert(0, p_pr)
                if p_pr.find(qn("w:keepNext")) is None:
                    p_pr.append(OxmlElement("w:keepNext"))


def _slice_body_elements(
    src_doc: Document,
    *,
    start_pattern: str | None = None,
    end_pattern: str | None = None,
) -> list:
    """Extract top-level paragraphs and tables from src_doc, optionally sliced between heading patterns."""
    elements = [
        child
        for child in src_doc.element.body
        if child.tag in {qn("w:p"), qn("w:tbl")}
    ]
    if not start_pattern and not end_pattern:
        return elements

    start_idx = 0
    if start_pattern:
        rx_start = re.compile(start_pattern)
        found = None
        for idx, el in enumerate(elements):
            if el.tag != qn("w:p"):
                continue
            txt = "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()
            if rx_start.search(txt):
                found = idx
                break
        if found is None:
            return []
        start_idx = found

    end_idx = len(elements)
    if end_pattern:
        rx_end = re.compile(end_pattern)
        for idx in range(start_idx + 1, len(elements)):
            if elements[idx].tag != qn("w:p"):
                continue
            txt = "".join(t.text or "" for t in elements[idx].iter(qn("w:t"))).strip()
            if txt and rx_end.search(txt):
                end_idx = idx
                break
    return elements[start_idx:end_idx]


def _make_styled_text_p_el(
    text: str,
    role: str,
    fmt: dict[str, Any],
    *,
    keep_next: bool = False,
    page_break_before: bool | None = None,
):
    p_el = OxmlElement("w:p")
    r_el = OxmlElement("w:r")
    t_el = OxmlElement("w:t")
    t_el.text = text
    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r_el.append(t_el)
    p_el.append(r_el)
    _apply_role_style_to_p_el(
        p_el,
        role,
        fmt,
        keep_next=keep_next,
        page_break_before=page_break_before,
    )
    return p_el


def _append_body_el(dst_doc: Document, el) -> None:
    sect_pr = dst_doc.element.body.find(qn("w:sectPr"))
    if sect_pr is not None:
        sect_pr.addprevious(el)
    else:
        dst_doc.element.body.append(el)


def _emit_title_el_with_trapezoid(
    out_els: list[Any],
    cloned_p,
    txt: str,
    role: str,
    fmt: dict[str, Any],
) -> None:
    """Apply clean punctuation, trapezoid wrapping, and optional ——subtitle split to a title paragraph."""
    main_txt, sub_txt = _split_title_and_subtitle(txt)
    main_clean = _clean_title_punctuation(main_txt)
    lines = _format_trapezoid_title_lines(main_clean)
    formatted_main = "\n".join(lines)
    if len(lines) > 1 or formatted_main != txt:
        _rewrite_paragraph_text(cloned_p, "")
        for old_r in list(cloned_p.findall(qn("w:r"))):
            cloned_p.remove(old_r)
        for li, ln in enumerate(lines):
            if li > 0:
                r_br = OxmlElement("w:r")
                r_br.append(OxmlElement("w:br"))
                cloned_p.append(r_br)
            r_ln = OxmlElement("w:r")
            t_ln = OxmlElement("w:t")
            t_ln.text = ln
            r_ln.append(t_ln)
            cloned_p.append(r_ln)
    _apply_role_style_to_p_el(cloned_p, role, fmt, keep_next=True)
    out_els.append(cloned_p)
    if sub_txt:
        sub_clean = f"——{_clean_title_punctuation(sub_txt[2:])}"
        out_els.append(_make_styled_text_p_el(sub_clean, "subtitle", fmt, keep_next=True))


def _split_p_with_inline_page_break(p_el) -> list[Any]:
    """Split a paragraph at <w:br w:type='page'/> when content exists on both sides."""
    if p_el.tag != qn("w:p"):
        return [p_el]
    if not any(br.get(qn("w:type")) == "page" for br in p_el.iter(qn("w:br"))):
        return [p_el]
    p_pr = p_el.find(qn("w:pPr"))
    segments: list[list[Any]] = [[]]
    for child in list(p_el):
        if child.tag == qn("w:pPr"):
            continue
        if child.tag == qn("w:r") and any(br.get(qn("w:type")) == "page" for br in child.findall(qn("w:br"))):
            r_pr = child.find(qn("w:rPr"))
            cur_r_children: list[Any] = []
            for rc in list(child):
                if rc.tag == qn("w:rPr"):
                    continue
                if rc.tag == qn("w:br") and rc.get(qn("w:type")) == "page":
                    if cur_r_children:
                        new_r = OxmlElement("w:r")
                        if r_pr is not None:
                            new_r.append(copy.deepcopy(r_pr))
                        for c in cur_r_children:
                            new_r.append(copy.deepcopy(c))
                        segments[-1].append(new_r)
                        cur_r_children = []
                    segments.append([])
                else:
                    cur_r_children.append(rc)
            if cur_r_children:
                new_r = OxmlElement("w:r")
                if r_pr is not None:
                    new_r.append(copy.deepcopy(r_pr))
                for c in cur_r_children:
                    new_r.append(copy.deepcopy(c))
                segments[-1].append(new_r)
        else:
            segments[-1].append(copy.deepcopy(child))
    non_empty_segs = [
        seg for seg in segments
        if any("".join(t.text or "" for t in ch.iter(qn("w:t"))).strip() or _has_media_el(ch) for ch in seg)
    ]
    if len(non_empty_segs) <= 1:
        for br in list(p_el.findall(".//" + qn("w:br"))):
            if br.get(qn("w:type")) == "page":
                parent = br.getparent()
                if parent is not None:
                    parent.remove(br)
        return [p_el]
    out_ps = []
    for idx_s, seg in enumerate(non_empty_segs):
        np = OxmlElement("w:p")
        if p_pr is not None:
            np_pr = copy.deepcopy(p_pr)
            if idx_s > 0:
                for pb in np_pr.findall(qn("w:pageBreakBefore")):
                    np_pr.remove(pb)
            np.append(np_pr)
        for ch in seg:
            np.append(ch)
        out_ps.append(np)
    return out_ps


def _import_docx_chapter(
    dst_doc: Document,
    src_path: Path,
    printable_w_twips: int,
    printable_h_twips: int,
    fmt: dict[str, Any],
    doc_pr_counter: list[int],
    *,
    start_pattern: str | None = None,
    end_pattern: str | None = None,
    skip_redundant_title: str = "",
    is_roster: bool = False,
    chapter_kind: str = "",
) -> None:
    with tempfile.TemporaryDirectory(prefix="imp-docx-") as tmp_dir:
        norm_docx = convert_to_docx(src_path, Path(tmp_dir) / "src.docx")
        src_doc = Document(str(norm_docx))
        elements = _slice_body_elements(
            src_doc, start_pattern=start_pattern, end_pattern=end_pattern
        )
        # Trim leading/trailing completely empty paragraphs
        while elements and elements[0].tag == qn("w:p") and not "".join(elements[0].itertext()).strip() and elements[0].find(".//" + qn("a:blip")) is None:
            elements.pop(0)
        while elements and elements[-1].tag == qn("w:p") and not "".join(elements[-1].itertext()).strip() and elements[-1].find(".//" + qn("a:blip")) is None:
            elements.pop()

        # Avoid duplicate bare chapter heading when the source Word file starts with exact same heading
        if elements and skip_redundant_title and elements[0].tag == qn("w:p"):
            first_txt = re.sub(r"\s+", "", "".join(t.text or "" for t in elements[0].iter(qn("w:t"))))
            clean_first = re.sub(r"^[一二三四五六七八九十]+[、.．]", "", first_txt)
            clean_target = re.sub(r"^[一二三四五六七八九十]+[、.．]", "", re.sub(r"\s+", "", skip_redundant_title))
            if clean_first and clean_target and (
                clean_first == clean_target
                or (clean_target == "组织机构" and clean_first in {"组织架构", "组织机构"})
                or (is_roster and (clean_target.startswith(clean_first) or f"（{clean_first}）" in clean_target))
            ):
                elements.pop(0)
                if is_roster and elements and elements[0].tag == qn("w:p"):
                    next_txt = re.sub(r"\s+", "", "".join(t.text or "" for t in elements[0].iter(qn("w:t"))))
                    if next_txt and f"（{next_txt}）" in clean_target:
                        elements.pop(0)

        # Split mixed <w:br w:type="page"/> paragraphs first, then collapse >1 consecutive blanks
        expanded_elements: list[Any] = []
        for raw_el in elements:
            expanded_elements.extend(_split_p_with_inline_page_break(raw_el))

        filtered: list[Any] = []
        prev_blank = False
        for el in expanded_elements:
            is_p = el.tag == qn("w:p")
            has_media = _has_media_el(el)
            txt = "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()
            if is_p and not txt and not has_media:
                if prev_blank:
                    continue
                prev_blank = True
            else:
                prev_blank = False
            filtered.append(el)

        # Keep short headings (even across spacer blank lines) glued to following tables or images
        keep_next_indices: set[int] = set()
        for idx in range(len(filtered) - 1, -1, -1):
            el = filtered[idx]
            if el.tag != qn("w:p"):
                continue
            has_media = _has_media_el(el)
            txt = "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()
            if has_media or len(txt) > 42:
                continue
            nxt_idx = idx + 1
            if nxt_idx < len(filtered):
                nxt = filtered[nxt_idx]
                nxt_media = nxt.tag == qn("w:p") and _has_media_el(nxt)
                if nxt.tag == qn("w:tbl") or nxt_media or nxt_idx in keep_next_indices:
                    keep_next_indices.add(idx)

        body_pt = float(fmt["fonts"]["body"].get("size_pt", 16.0))
        line_pt = float(fmt["fonts"]["body"].get("line_spacing_pt", 30.0))
        out_els: list[Any] = []

        if is_roster:
            for idx, el in enumerate(filtered):
                cloned = _remap_and_sanitize_element(
                    el, src_doc.part, dst_doc, doc_pr_counter, printable_w_twips, printable_h_twips
                )
                if el.tag == qn("w:p"):
                    _style_imported_paragraph(cloned, fmt, is_roster=True)
                    if idx in keep_next_indices:
                        p_pr = cloned.find(qn("w:pPr"))
                        if p_pr is None:
                            p_pr = OxmlElement("w:pPr")
                            cloned.insert(0, p_pr)
                        if p_pr.find(qn("w:keepNext")) is None:
                            p_pr.append(OxmlElement("w:keepNext"))
                out_els.append(cloned)
            for oe in out_els:
                _append_body_el(dst_doc, oe)
            return

        def _is_blank_p(el_cand) -> bool:
            if el_cand is None or el_cand.tag != qn("w:p"):
                return False
            return not "".join(t.text or "" for t in el_cand.iter(qn("w:t"))).strip() and not _has_media_el(el_cand)

        def _is_doc_title_candidate(el_cand, is_first: bool) -> bool:
            if el_cand is None or el_cand.tag != qn("w:p") or _has_media_el(el_cand):
                return False
            t = "".join(t.text or "" for t in el_cand.iter(qn("w:t"))).strip()
            if not t or len(t) > 44:
                return False
            _, sk = _detect_and_normalize_structure(t)
            if sk != "body" or _NO_INDENT_ROLE_RE.match(t):
                return False
            if t.endswith(("。", "；")) and "——" not in t:
                return False
            jc = _xml_jc(el_cand)
            if jc == "center":
                return True
            if is_first and re.search(r"(竞赛规程|总体方案|活动方案|补充通知|组织机构|组织架构|承诺书|须知|说明|指南)$", t):
                return True
            return False

        in_top_title_zone = True
        expecting_att_title = False
        att_title_lines = 0
        last_role = ""
        i = 0
        n_el = len(filtered)
        while i < n_el:
            el = filtered[i]
            if el.tag == qn("w:tbl"):
                in_top_title_zone = False
                expecting_att_title = False
                att_title_lines = 0
                cloned = _remap_and_sanitize_element(
                    el, src_doc.part, dst_doc, doc_pr_counter, printable_w_twips, printable_h_twips
                )
                _style_imported_table(cloned, printable_w_twips, fmt, is_roster=False)
                out_els.append(cloned)
                last_role = "table"
                i += 1
                continue

            if el.tag != qn("w:p"):
                cloned = _remap_and_sanitize_element(
                    el, src_doc.part, dst_doc, doc_pr_counter, printable_w_twips, printable_h_twips
                )
                out_els.append(cloned)
                i += 1
                continue

            cloned = _remap_and_sanitize_element(
                el, src_doc.part, dst_doc, doc_pr_counter, printable_w_twips, printable_h_twips
            )
            has_media = _has_media_el(cloned)
            txt = "".join(t.text or "" for t in cloned.iter(qn("w:t"))).strip()

            # Blank paragraph normalization
            if not txt and not has_media:
                p_pr0 = cloned.find(qn("w:pPr"))
                if p_pr0 is not None:
                    for pb in p_pr0.findall(qn("w:pageBreakBefore")):
                        p_pr0.remove(pb)
                for br in list(cloned.findall(".//" + qn("w:br"))):
                    if br.get(qn("w:type")) == "page":
                        parent = br.getparent()
                        if parent is not None:
                            parent.remove(br)
                role_b = "blank_10_5" if last_role in {"title", "attachment_title", "subtitle"} else "blank_line"
                kn_b = (i in keep_next_indices) or last_role in {"attachment_label", "title", "attachment_title", "subtitle"}
                _style_imported_paragraph(cloned, fmt, role_override=role_b, keep_next=kn_b)
                out_els.append(cloned)
                last_role = role_b
                if last_role == "blank_line" and not expecting_att_title:
                    in_top_title_zone = False
                i += 1
                continue

            norm_txt, struct_kind = _detect_and_normalize_structure(txt)

            # Top-of-chapter document title lines (Rule 1)
            if in_top_title_zone and _is_doc_title_candidate(el, i == 0):
                _emit_title_el_with_trapezoid(out_els, cloned, norm_txt, "title", fmt)
                last_role = "title"
                nxt = filtered[i + 1] if (i + 1) < n_el else None
                if not _is_doc_title_candidate(nxt, False):
                    in_top_title_zone = False
                    if nxt is not None and not _is_blank_p(nxt):
                        out_els.append(_make_blank_spacer_p_el(10.5, line_pt=13.0, keep_next=True))
                        last_role = "blank_10_5"
                i += 1
                continue

            in_top_title_zone = False

            # Rule 4: Attachment page marker ("附件" / "附件1") -> new page, Line 1 label, Line 2 blank, Line 3 title
            if struct_kind == "attachment_label":
                expecting_att_title = False
                att_title_lines = 0
                # Drop trailing blank spacer before pageBreakBefore so no stray blank page occurs
                while out_els and _is_blank_p(out_els[-1]):
                    out_els.pop()
                for br in list(cloned.findall(".//" + qn("w:br"))):
                    if br.get(qn("w:type")) == "page":
                        parent = br.getparent()
                        if parent is not None:
                            parent.remove(br)
                if norm_txt != txt:
                    _rewrite_paragraph_text(cloned, norm_txt)
                need_pb = len(out_els) > 0
                _style_imported_paragraph(
                    cloned,
                    fmt,
                    role_override="attachment_label",
                    keep_next=True,
                    page_break_before=need_pb,
                )
                out_els.append(cloned)
                last_role = "attachment_label"
                nxt = filtered[i + 1] if (i + 1) < n_el else None
                if nxt is not None and not _is_blank_p(nxt):
                    # Ensure Line 2 blank line so Attachment Title lands on Line 3
                    out_els.append(_make_blank_spacer_p_el(body_pt, line_pt=line_pt, keep_next=True))
                    last_role = "blank_line"
                expecting_att_title = True
                i += 1
                continue

            # Rule 4: Line 3 centered attachment title following attachment_label + Line 2 blank
            if expecting_att_title:
                if (
                    not has_media
                    and len(txt) <= 44
                    and struct_kind == "body"
                    and not _NO_INDENT_ROLE_RE.match(txt)
                    and att_title_lines < 2
                    and (att_title_lines == 0 or _xml_jc(el) == "center")
                ):
                    _emit_title_el_with_trapezoid(out_els, cloned, norm_txt, "attachment_title", fmt)
                    att_title_lines += 1
                    last_role = "attachment_title"
                    nxt = filtered[i + 1] if (i + 1) < n_el else None
                    nxt_is_more_title = (
                        nxt is not None
                        and nxt.tag == qn("w:p")
                        and att_title_lines < 2
                        and _xml_jc(nxt) == "center"
                        and not _is_blank_p(nxt)
                        and not _NO_INDENT_ROLE_RE.match("".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip())
                    )
                    if not nxt_is_more_title:
                        expecting_att_title = False
                        nxt_txt = "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip() if nxt is not None else ""
                        if (
                            nxt is not None
                            and nxt.tag == qn("w:p")
                            and not _is_blank_p(nxt)
                            and not _has_media_el(nxt)
                            and not _NO_INDENT_ROLE_RE.match(nxt_txt)
                            and not re.search(r"[：:]\s{3,}\S", nxt_txt)
                        ):
                            out_els.append(_make_blank_spacer_p_el(10.5, line_pt=13.0, keep_next=True))
                            last_role = "blank_10_5"
                    i += 1
                    continue
                expecting_att_title = False
                att_title_lines = 0

            # Rule 3: "附件说明" group ("附件：..." + optional subsequent "2．..." paragraphs)
            if struct_kind == "attachment_desc":
                if out_els and not _is_blank_p(out_els[-1]):
                    out_els.append(_make_blank_spacer_p_el(body_pt, line_pt=line_pt))
                parsed_items = _split_attachment_desc_items(txt)
                # Look ahead for separately-paragraph-broken sequential "2．..." attachment list lines
                j = i + 1
                expected_num = len(parsed_items) + 1 if re.match(r"^附件：1．", parsed_items[0][0]) else None
                while (
                    j < n_el
                    and expected_num is not None
                    and filtered[j].tag == qn("w:p")
                    and not _has_media_el(filtered[j])
                ):
                    jt = "".join(t.text or "" for t in filtered[j].iter(qn("w:t"))).strip()
                    m_next = re.match(r"^(\d{1,2})\s*[.．、，,]\s*([^\d\s.%％].*)$", jt)
                    if m_next and int(m_next.group(1)) == expected_num and "。" not in jt[:-1]:
                        clean_item = f"{m_next.group(1)}．{re.sub(r'[。；;.,，、]+$', '', m_next.group(2).strip())}"
                        parsed_items.append((clean_item, False))
                        expected_num += 1
                        j += 1
                    else:
                        break
                is_multi = len(parsed_items) > 1 or bool(re.match(r"^附件：\d+．", parsed_items[0][0]))
                for idx_it, (it_txt, is_first) in enumerate(parsed_items):
                    role_desc = (
                        ("attachment_desc_multi_first" if is_first else "attachment_desc_multi_next")
                        if is_multi
                        else "attachment_desc_single"
                    )
                    if idx_it == 0:
                        _rewrite_paragraph_text(cloned, it_txt)
                        _style_imported_paragraph(cloned, fmt, role_override=role_desc)
                        out_els.append(cloned)
                    else:
                        out_els.append(_make_styled_text_p_el(it_txt, role_desc, fmt))
                last_role = "attachment_desc"
                i = j
                continue

            # Regular body or Rule 2 heading paragraph
            _style_imported_paragraph(
                cloned,
                fmt,
                is_roster=False,
                keep_next=(i in keep_next_indices),
                chapter_kind=chapter_kind,
            )
            out_els.append(cloned)
            last_role = struct_kind
            i += 1

        for oe in out_els:
            _append_body_el(dst_doc, oe)


def _import_pdf_chapter(
    dst_doc: Document,
    pdf_path: Path,
    printable_w_twips: int,
    printable_h_twips: int,
) -> None:
    with tempfile.TemporaryDirectory(prefix="imp-pdf-") as tmp_dir:
        tmp = Path(tmp_dir)
        with pymupdf.open(pdf_path) as pdf_doc:
            for idx in range(pdf_doc.page_count):
                pix = pdf_doc[idx].get_pixmap(matrix=pymupdf.Matrix(2.4, 2.4))
                img_file = tmp / f"page_{idx + 1:03d}.png"
                pix.save(str(img_file))
                _add_scaled_image_paragraph(dst_doc, img_file, printable_w_twips, printable_h_twips)


def _add_text_blocks(dst_doc: Document, blocks: list[str], fmt: dict[str, Any]) -> None:
    lines: list[str] = []
    for blk in blocks:
        lines.extend(str(blk).splitlines())
    body_pt = float(fmt["fonts"]["body"].get("size_pt", 16.0))
    line_pt = float(fmt["fonts"]["body"].get("line_spacing_pt", 30.0))

    expecting_att_title = False
    i = 0
    while i < len(lines):
        raw = lines[i].strip()
        if not raw:
            i += 1
            continue
        norm, role = _detect_and_normalize_structure(raw)

        if role == "attachment_label":
            _append_body_el(
                dst_doc,
                _make_styled_text_p_el(norm, "attachment_label", fmt, keep_next=True, page_break_before=True),
            )
            _append_body_el(dst_doc, _make_blank_spacer_p_el(body_pt, line_pt=line_pt, keep_next=True))
            expecting_att_title = True
            i += 1
            continue

        if expecting_att_title and role == "body" and len(norm) <= 44 and not _NO_INDENT_ROLE_RE.match(norm):
            main_t, sub_t = _split_title_and_subtitle(_clean_title_punctuation(norm))
            t_lines = _format_trapezoid_title_lines(main_t)
            _append_body_el(dst_doc, _make_styled_text_p_el("\n".join(t_lines), "attachment_title", fmt, keep_next=True))
            if sub_t:
                _append_body_el(dst_doc, _make_styled_text_p_el(sub_t, "subtitle", fmt, keep_next=True))
            _append_body_el(dst_doc, _make_blank_spacer_p_el(10.5, line_pt=13.0, keep_next=True))
            expecting_att_title = False
            i += 1
            continue
        expecting_att_title = False

        if role == "attachment_desc":
            _append_body_el(dst_doc, _make_blank_spacer_p_el(body_pt, line_pt=line_pt))
            items = _split_attachment_desc_items(raw)
            j = i + 1
            while j < len(lines):
                nxt = lines[j].strip()
                m_n = re.match(r"^(\d{1,2})\s*[.．、，,]\s*([^\d\s.%％].*)$", nxt)
                if m_n:
                    items.append((f"{m_n.group(1)}．{re.sub(r'[。；;.,，、]+$', '', m_n.group(2).strip())}", False))
                    j += 1
                elif not nxt:
                    j += 1
                else:
                    break
            is_multi = len(items) > 1 or bool(re.match(r"^附件：\d+．", items[0][0]))
            for it_txt, is_first in items:
                r_name = (
                    ("attachment_desc_multi_first" if is_first else "attachment_desc_multi_next")
                    if is_multi
                    else "attachment_desc_single"
                )
                _append_body_el(dst_doc, _make_styled_text_p_el(it_txt, r_name, fmt))
            i = j
            continue

        _append_body_el(dst_doc, _make_styled_text_p_el(norm, role, fmt))
        i += 1


def _resolve_chapter_landscape(chap: dict[str, Any]) -> bool:
    flag = chap.get("landscape", False)
    if flag is True:
        return True
    if flag != "auto":
        return False
    for src in chap.get("sources") or []:
        if src.get("type") in {"xlsx", "xls", "et"} and src.get("path"):
            p = Path(src["path"])
            if p.is_file():
                cols = xlsx_sheet_dimensions(p, src.get("sheets"))
                if cols >= 10:
                    return True
    return False


def assemble_book(
    plan: dict[str, Any],
    template: dict[str, Any],
    out_docx: Path,
    out_pdf: Path,
    *,
    watermark_override: Path | None = None,
) -> dict[str, Any]:
    """Assemble a complete multi-chapter Event Order Book (.docx + TOC-synced .pdf) from a BuildPlan + Template."""
    out_docx = Path(out_docx).resolve()
    out_pdf = Path(out_pdf).resolve()
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    out_pdf.parent.mkdir(parents=True, exist_ok=True)

    fmt = template.get("formatting") or {}
    page_cfg = fmt.get("page", {})
    w_mm = float(page_cfg.get("width_mm", 210.0))
    h_mm = float(page_cfg.get("height_mm", 297.0))

    doc = Document()
    # Set sane Normal style baseline (0pt paragraph spacing so 30pt exact lines yield exact 22 lines/page)
    normal_style = doc.styles["Normal"]
    body_cfg = fmt.get("fonts", {}).get("body", {})
    body_font_name = str(body_cfg.get("name", "方正仿宋_GBK"))
    body_font_pt = float(body_cfg.get("size_pt", 16.0))
    normal_style.font.name = body_font_name
    normal_style.font.size = Pt(body_font_pt)
    normal_rpr = normal_style._element.get_or_add_rPr()
    normal_rf = normal_rpr.find(qn("w:rFonts"))
    if normal_rf is None:
        normal_rf = OxmlElement("w:rFonts")
        normal_rpr.insert(0, normal_rf)
    for attr in ("ascii", "eastAsia", "hAnsi", "cs"):
        normal_rf.set(qn(f"w:{attr}"), body_font_name)
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(0)
    normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    normal_style.paragraph_format.line_spacing = Pt(float(body_cfg.get("line_spacing_pt", 30.0)))

    pn_cfg = fmt.get("page_numbers", {})
    pn_align = str(pn_cfg.get("align", "odd_right_even_left")).lower()
    if pn_cfg.get("odd_even", pn_align == "odd_right_even_left"):
        doc.settings.odd_and_even_pages_header_footer = True
    enable_update_fields_on_open(doc)

    doc_pr_counter = [1000]
    bookmark_id = [10]

    # Resolve optional watermark path (user project底图/template watermark)
    wm_path = watermark_override
    if wm_path is None and fmt.get("watermark", {}).get("enabled") and fmt["watermark"].get("asset"):
        from templates_store import template_asset_path

        wm_path = template_asset_path(template["id"], fmt["watermark"]["asset"])

    # 1. Cover Section (Section 0)
    cover_info = plan.get("cover") or {}
    cover_path = Path(cover_info["path"]) if cover_info.get("path") else None
    sec0 = doc.sections[0]
    _clear_section_hf(sec0)

    has_cover_page = False
    if cover_path and cover_path.is_file():
        mode = cover_info.get("mode", "image")
        with tempfile.TemporaryDirectory(prefix="cover-") as tmp_c:
            img_for_cover = cover_path
            if mode == "background":
                img_for_cover = generate_cover_from_background(
                    cover_path,
                    Path(tmp_c) / "generated_cover.png",
                    plan.get("event_title") or "体育赛事",
                    subtitle=plan.get("subtitle") or "秩  序  册",
                    footer_lines=plan.get("footer_lines") or [],
                )
            elif cover_path.suffix.lower() == ".pdf":
                with pymupdf.open(cover_path) as pdfc:
                    pix = pdfc[0].get_pixmap(matrix=pymupdf.Matrix(2.4, 2.4))
                    img_for_cover = Path(tmp_c) / "cover_pdf_p1.png"
                    pix.save(str(img_for_cover))
            _configure_section(sec0, fmt, landscape=False, zero_margins=True)
            _insert_fullpage_image(doc, img_for_cover, w_mm, h_mm, doc_pr_counter)
            has_cover_page = True
    else:
        # Clean typographic cover if no cover image uploaded
        _configure_section(sec0, fmt, landscape=False, zero_margins=False)
        for _ in range(3):
            doc.add_paragraph()
        tp = doc.add_paragraph()
        tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tp.paragraph_format.space_after = Pt(24)
        tr = tp.add_run(plan.get("event_title") or "体育赛事")
        _set_run_font(tr, str(fmt.get("fonts", {}).get("title", {}).get("name", "方正小标宋_GBK")), 26.0, bold=True)
        sp = doc.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp.paragraph_format.space_after = Pt(120)
        sr = sp.add_run(plan.get("subtitle") or "秩  序  册")
        _set_run_font(sr, str(fmt.get("fonts", {}).get("title", {}).get("name", "方正小标宋_GBK")), 34.0, bold=True)
        for fl in (plan.get("footer_lines") or [])[:3]:
            fp = doc.add_paragraph()
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            fr = fp.add_run(fl)
            _set_run_font(fr, str(fmt.get("fonts", {}).get("heading1", {}).get("name", "方正黑体_GBK")), 16.0)
        has_cover_page = True

    # 2. Gather renderable body chapters and build TOC metadata first
    render_chapters: list[dict[str, Any]] = []
    chap_counter = 0
    for idx, ch in enumerate(plan.get("chapters") or []):
        kind = ch.get("kind", "richtext")
        if kind in {"cover", "toc", "back_cover"}:
            continue
        if kind == "notes":
            render_chapters.append(
                {
                    **ch,
                    "anchor": "_toc_notes",
                    "display_title": ch.get("title") or "记    事",
                    "toc_visible": False,
                    "is_landscape": False,
                }
            )
            continue
        sources = ch.get("sources") or []
        if not sources:
            continue
        chap_counter += 1
        raw_title = str(ch.get("title") or f"第{_cn_index(chap_counter)}章").strip()
        anchor = f"_toc_{chap_counter}"
        display_title = (
            raw_title
            if re.match(r"^[一二三四五六七八九十百]+[、.．]", raw_title)
            else f"{_cn_index(chap_counter)}、{raw_title}"
        )
        toc_label = raw_title if ch.get("toc_label") is None else str(ch["toc_label"])
        render_chapters.append(
            {
                **ch,
                "anchor": anchor,
                "display_title": display_title,
                "toc_label": toc_label,
                "toc_visible": bool(ch.get("toc_visible", True)),
                "is_landscape": _resolve_chapter_landscape(ch),
            }
        )

    toc_targets = [
        {
            "anchor": ch["anchor"],
            "title": ch["display_title"],
            "toc_label": ch["toc_label"],
        }
        for ch in render_chapters
        if ch.get("toc_visible")
    ]

    # Insert TOC section cleanly (Section 1) without page numbers or watermark
    if toc_targets:
        sec_toc = doc.add_section(WD_SECTION.NEW_PAGE)
        toc_w_twips, _ = _configure_section(sec_toc, fmt, landscape=False)
        _clear_section_hf(sec_toc)
        _add_toc_block(doc, toc_targets, toc_w_twips, fmt)

    # 3. Body Sections (Page numbering restarts at 1 on the first body chapter)
    current_landscape: bool | None = None
    body_started = False
    for ch_idx, ch in enumerate(render_chapters):
        want_land = bool(ch.get("is_landscape", False))
        need_page_break = False
        if not body_started or current_landscape != want_land:
            sec = doc.add_section(WD_SECTION.NEW_PAGE)
            pw_twips, ph_twips = _configure_section(sec, fmt, landscape=want_land)
            restart = 1 if not body_started else None
            _apply_section_footer_page_number(sec, fmt, restart_at=restart)
            _apply_section_watermark(
                sec,
                wm_path,
                max(w_mm, h_mm) if want_land else min(w_mm, h_mm),
                min(w_mm, h_mm) if want_land else max(w_mm, h_mm),
                doc_pr_counter,
            )
            current_landscape = want_land
            body_started = True
        else:
            # Start chapter cleanly on a new page via pageBreakBefore (avoids blank trailing pages on full 22-line pages)
            need_page_break = True

        kind = ch.get("kind", "richtext")
        if kind == "notes":
            np = doc.add_paragraph()
            np.alignment = WD_ALIGN_PARAGRAPH.CENTER
            np.paragraph_format.space_before = Pt(0)
            np.paragraph_format.space_after = Pt(0)
            if need_page_break:
                np.paragraph_format.page_break_before = True
            nr = np.add_run("记    事")
            _set_run_font(
                nr,
                str(fmt.get("fonts", {}).get("title", {}).get("name", "方正小标宋_GBK")),
                22.0,
                bold=False,
            )
            continue

        bookmark_id[0] += 1
        _add_bookmarked_chapter_heading(
            doc,
            ch["display_title"],
            ch["anchor"],
            bookmark_id[0],
            fmt,
            centered=True,
            page_break_before=need_page_break,
        )

        is_roster_chapter = kind == "roster"
        for src in ch.get("sources") or []:
            stype = str(src.get("type") or "").lower()
            if stype == "text":
                _add_text_blocks(doc, src.get("text_blocks") or [], fmt)
                continue
            src_path = Path(src["path"]) if src.get("path") else None
            if not src_path or not src_path.is_file():
                continue
            caption = str(src.get("caption") or "").strip()
            if caption and stype in {"image", "png", "jpg", "jpeg", "webp"}:
                cp = doc.add_paragraph()
                cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                cp.paragraph_format.space_before = Pt(6)
                cp.paragraph_format.space_after = Pt(4)
                cp.paragraph_format.first_line_indent = Pt(0)
                cp.paragraph_format.keep_with_next = True
                cr = cp.add_run(caption)
                _set_run_font(cr, str(fmt["fonts"]["heading2"].get("name", "方正楷体_GBK")), 16.0, bold=False)

            if stype in {"docx", "doc", "wps"}:
                _import_docx_chapter(
                    doc,
                    src_path,
                    pw_twips,
                    ph_twips,
                    fmt,
                    doc_pr_counter,
                    start_pattern=src.get("start_pattern"),
                    end_pattern=src.get("end_pattern"),
                    skip_redundant_title=ch.get("title") or "",
                    is_roster=is_roster_chapter,
                    chapter_kind=kind,
                )
            elif stype in {"xlsx", "xls", "et"}:
                _import_xlsx_chapter(
                    doc,
                    src_path,
                    pw_twips,
                    fmt,
                    sheet_names=src.get("sheets"),
                    chapter_title=ch.get("title") or "",
                )
            elif stype == "pdf":
                _import_pdf_chapter(doc, src_path, pw_twips, ph_twips)
            elif stype in {"image", "png", "jpg", "jpeg", "webp"}:
                _add_scaled_image_paragraph(doc, src_path, pw_twips, ph_twips)

    # 4. Back Cover (full-bleed trailing section without page numbers or watermark)
    back_info = plan.get("back_cover") or {}
    back_path = Path(back_info["path"]) if back_info.get("path") else None
    if back_path and back_path.is_file():
        sec_back = doc.add_section(WD_SECTION.NEW_PAGE)
        _configure_section(sec_back, fmt, landscape=False, zero_margins=True)
        _clear_section_hf(sec_back)
        with tempfile.TemporaryDirectory(prefix="back-") as tmp_b:
            img_for_back = back_path
            if back_path.suffix.lower() == ".pdf":
                with pymupdf.open(back_path) as pdfb:
                    pix = pdfb[-1].get_pixmap(matrix=pymupdf.Matrix(2.4, 2.4))
                    img_for_back = Path(tmp_b) / "back_pdf.png"
                    pix.save(str(img_for_back))
            _insert_fullpage_image(doc, img_for_back, w_mm, h_mm, doc_pr_counter)

    tmp_docx = out_docx.with_suffix(".docx.tmp")
    doc.save(str(tmp_docx))
    tmp_docx.replace(out_docx)

    _, resolved_toc = sync_toc_and_export_pdf(
        out_docx,
        out_pdf,
        toc_targets=toc_targets,
        timeout=180.0,
    )

    with pymupdf.open(out_pdf) as pdf_final:
        total_pdf_pages = pdf_final.page_count

    return {
        "docx_path": str(out_docx),
        "pdf_path": str(out_pdf),
        "pdf_pages": total_pdf_pages,
        "has_cover": has_cover_page,
        "has_back_cover": bool(back_path and back_path.is_file()),
        "toc": resolved_toc,
        "chapters": [
            {
                "id": ch.get("id"),
                "title": ch["display_title"],
                "kind": ch.get("kind"),
                "landscape": bool(ch.get("is_landscape")),
                "toc_visible": bool(ch.get("toc_visible")),
            }
            for ch in render_chapters
        ],
    }
