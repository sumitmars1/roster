from __future__ import annotations

import json
import re
import tempfile
import urllib.request
from collections import Counter
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn

import config
from book_pdf import convert_to_docx
from presets import PRESETS
from templates_store import _default_official_formatting, save_custom_template

TWIPS_PER_MM = 56.6929133858
HALF_PT_PER_PT = 2.0

_CN_NUM = r"[一二三四五六七八九十百]+"

# Map common chapter titles to standardized template chapter kinds
_CHAPTER_KIND_RULES: list[tuple[str, str, str]] = [
    (r"目录", "toc", "目录"),
    (r"竞赛规程|比赛规程|赛事规程|总体方案|活动方案|方案$", "richtext", "竞赛规程"),
    (r"补充通知|领队会|会议纪要|注意事项|安全须知|应急预案", "richtext", "补充通知"),
    (r"组织机构|组织架构|组委会|裁判名单|裁判组", "organization", "组织机构"),
    (r"代表队名单|参赛名单|运动员名单|参赛单位|队伍名单|人员名单", "roster", "各代表队名单"),
    (r"分组|签位|抽签", "grouping", "分组情况"),
    (r"日程|赛程|轮转|竞赛安排|时间安排", "schedule", "竞赛日程表"),
    (r"积分表|成绩公告|成绩统计", "standings", "比赛积分表"),
    (r"对阵|淘汰赛|决赛图|名次", "brackets", "淘汰赛对阵表"),
    (r"指南|须知|交通|食宿|报到|站位|平面|示意图|赞助|简介|大巴|联系", "richtext", "参赛指南"),
    (r"记\s*事|备\s*忘", "notes", "记事"),
]


_LATIN_ONLY_FONTS = {
    "times new roman",
    "arial",
    "calibri",
    "cambria",
    "consolas",
    "courier new",
    "helvetica",
    "symbol",
    "wingdings",
}


def _length_to_mm(val: Any, default_mm: float, *, min_mm: float = 5.0, max_mm: float = 450.0) -> float:
    if not val:
        return default_mm
    mm = float(val.mm) if hasattr(val, "mm") else (float(val) / 36000.0)
    if min_mm <= mm <= max_mm:
        return round(mm, 1)
    return default_mm


def _is_cjk_capable_font(name: str | None) -> bool:
    return bool(name and name.strip().lower() not in _LATIN_ONLY_FONTS)


def _paragraph_jc(p) -> str:
    """Safely read paragraph horizontal alignment from XML without raising on 'start'/'end'."""
    p_pr = p._p.find(qn("w:pPr"))
    if p_pr is not None:
        jc_el = p_pr.find(qn("w:jc"))
        if jc_el is not None:
            val = (jc_el.get(qn("w:val")) or "").lower()
            if val in {"start", "left"}:
                return "left"
            if val in {"end", "right"}:
                return "right"
            if val:
                return val
    return ""


def _extract_run_font(r_pr) -> tuple[str | None, float | None]:
    if r_pr is None:
        return None, None
    font_name = None
    rfonts = r_pr.find(qn("w:rFonts"))
    if rfonts is not None:
        font_name = (
            rfonts.get(qn("w:eastAsia"))
            or rfonts.get(qn("w:ascii"))
            or rfonts.get(qn("w:hAnsi"))
        )
    sz = r_pr.find(qn("w:sz"))
    size_pt = None
    if sz is not None:
        raw = sz.get(qn("w:val"))
        if raw and raw.isdigit():
            size_pt = round(int(raw) / HALF_PT_PER_PT, 1)
    return font_name, size_pt


def _paragraph_font(p) -> tuple[str | None, float | None, bool]:
    names: list[str] = []
    sizes: list[float] = []
    bold = False
    for run in p._p.findall(qn("w:r")):
        r_pr = run.find(qn("w:rPr"))
        fn, sz = _extract_run_font(r_pr)
        if fn:
            names.append(fn)
        if sz:
            sizes.append(sz)
        if r_pr is not None and r_pr.find(qn("w:b")) is not None:
            b_val = r_pr.find(qn("w:b")).get(qn("w:val"))
            if b_val not in {"0", "false"}:
                bold = True
    if not names or not sizes:
        p_pr = p._p.find(qn("w:pPr"))
        if p_pr is not None:
            fn, sz = _extract_run_font(p_pr.find(qn("w:rPr")))
            if fn and not names:
                names.append(fn)
            if sz and not sizes:
                sizes.append(sz)
    style = getattr(p, "style", None)
    if style is not None:
        style_el = getattr(style, "_element", None)
        if style_el is not None:
            fn, sz = _extract_run_font(style_el.find(qn("w:rPr")))
            if fn and not names:
                names.append(fn)
            if sz and not sizes:
                sizes.append(sz)
    return (names[0] if names else None, sizes[0] if sizes else None, bold)


def _strip_heading_prefix(text: str) -> str:
    cleaned = re.sub(r"\s+", " ", (text or "").strip())
    cleaned = re.sub(rf"^{_CN_NUM}[、.．]\s*", "", cleaned)
    cleaned = re.sub(r"^\d{1,2}[.、．]\s*", "", cleaned)
    cleaned = re.sub(r"^第[一二三四五六七八九十\d]+[章部分篇节]\s*", "", cleaned)
    return cleaned.strip()


def _sanitize_chapter_title(raw_title: str, kind: str) -> str:
    """Strip historical event names/years/group specifics while keeping reusable chapter title."""
    title = _strip_heading_prefix(raw_title)
    # Remove trailing or leading group qualifiers like "（部门男子组）" / "部门男子组..."
    title = re.sub(r"[（(][^()（）]*组[）)]$", "", title).strip()
    title = re.sub(r"^.*?(?:男子组|女子组|男女混合组|混合组|公开组|青年组|少年组|甲组|乙组|丙组)\s*", "", title).strip()
    # Strip specific event prefixes like "开幕式暨趣味运动会..." or regional event names before standard chapter nouns
    title = re.sub(r"^.*?暨[^，。、]*(?=(总体方案|整体流程|参赛单位|竞赛安排|日程|规程))", "", title).strip()
    title = re.sub(r"^.*?(趣味运动会|篮球赛|足球赛|运动会)(?=(参赛单位|竞赛安排|整体流程|总体方案|竞赛规程|竞赛日程|各代表队名单))", "", title).strip()
    if re.search(r"20\d{2}|第[一二三四五六七八九十\d]+届", title):
        for pattern, _k, generic in _CHAPTER_KIND_RULES:
            if re.search(pattern, title):
                return generic
    if not title:
        for _pattern, k, generic in _CHAPTER_KIND_RULES:
            if k == kind:
                return generic
        return "附加章节"
    return title[:40]


def _classify_chapter_kind(title: str) -> tuple[str, str]:
    clean = _strip_heading_prefix(title)
    for pattern, kind, generic in _CHAPTER_KIND_RULES:
        if re.search(pattern, clean):
            return kind, _sanitize_chapter_title(clean, kind) or generic
    return "richtext", _sanitize_chapter_title(clean, "richtext")


def _extract_watermark_asset(doc: Document, out_dir: Path) -> Path | None:
    """Extract fixed header background/watermark image (e.g. '底图' / WordPictureWatermark) cleanly."""
    seen_rids: set[str] = set()
    for section in doc.sections:
        for hdr_attr in ("header", "first_page_header", "even_page_header"):
            hdr = getattr(section, hdr_attr, None)
            if hdr is None or hdr.is_linked_to_previous:
                continue
            hdr_xml = hdr._element.xml
            if not any(
                marker in hdr_xml
                for marker in ("底图", "Watermark", "v:imagedata", "a:blip")
            ):
                continue
            part = hdr.part
            for blip in hdr._element.iter(qn("a:blip")):
                rid = blip.get(qn("r:embed"))
                if rid and rid not in seen_rids and rid in part.rels:
                    seen_rids.add(rid)
                    rel = part.rels[rid]
                    target = rel.target_part
                    ext = Path(str(target.partname)).suffix or ".png"
                    dest = out_dir / f"extracted_watermark{ext}"
                    dest.write_bytes(target.blob)
                    return dest
            V_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"
            for v_img in hdr._element.iter(V_IMAGEDATA):
                rid = v_img.get(qn("r:id"))
                if rid and rid not in seen_rids and rid in part.rels:
                    seen_rids.add(rid)
                    target = part.rels[rid].target_part
                    ext = Path(str(target.partname)).suffix or ".png"
                    dest = out_dir / f"extracted_watermark{ext}"
                    dest.write_bytes(target.blob)
                    return dest
    return None


def _detect_page_number_style(doc: Document) -> dict[str, Any]:
    style = dict(_default_official_formatting()["page_numbers"])
    odd_even_enabled = bool(getattr(doc.settings, "odd_and_even_pages_header_footer", False))
    for sec in doc.sections:
        odd_align = None
        even_align = None
        found_ftr = None
        for attr in ("footer", "even_page_footer", "first_page_footer"):
            ftr = getattr(sec, attr, None)
            if ftr is None or "PAGE" not in ftr._element.xml:
                continue
            if found_ftr is None:
                found_ftr = ftr
            for p in ftr.paragraphs:
                jc = _paragraph_jc(p)
                if attr == "footer":
                    odd_align = jc or "center"
                elif attr == "even_page_footer":
                    even_align = jc or "center"
                break
        if found_ftr is None:
            continue

        text = "".join(t.text or "" for t in found_ftr._element.iter(qn("w:t")))
        compact_text = re.sub(r"\s+", "", text)
        if "第" in text and "页" in text:
            style["style"] = "di_ye"
            style["prefix"] = "第 "
            style["suffix"] = " 页"
        elif any(ch in text for ch in ("—", "–")):
            style["style"] = "em_dash"
            style["prefix"] = "— "
            style["suffix"] = " —"
        elif any(ch in text for ch in ("-", "－")):
            style["style"] = "hyphen"
            has_inner_space = bool(re.search(r"-\s+\d|\d\s+-", text))
            style["prefix"] = "- " if has_inner_space and not re.search(r"-\d+-", compact_text) else "-"
            style["suffix"] = " -" if has_inner_space and not re.search(r"-\d+-", compact_text) else "-"
        else:
            style["style"] = "plain"
            style["prefix"] = ""
            style["suffix"] = ""

        for p in found_ftr.paragraphs:
            fn, sz, _ = _paragraph_font(p)
            if fn and _is_cjk_capable_font(fn):
                style["font"] = fn
            if sz and sz >= 10.0:
                style["size_pt"] = sz
            break

        if odd_even_enabled and odd_align == "right" and even_align == "left":
            style["align"] = "odd_right_even_left"
            style["odd_even"] = True
            style.setdefault("indent_chars", 1.0)
        elif odd_align in {"left", "right", "center"}:
            style["align"] = odd_align
            style["odd_even"] = False
        return style
    return style


def _detect_formatting(doc: Document) -> dict[str, Any]:
    fmt = _default_official_formatting()
    portrait_secs = [s for s in doc.sections if (s.page_width or 0) <= (s.page_height or 1)]
    candidates = portrait_secs[1:] if len(portrait_secs) > 1 else portrait_secs
    if candidates:
        geom_counts: Counter[tuple[float, float, float, float, float, float]] = Counter()
        sec_by_geom: dict[tuple[float, float, float, float, float, float], Any] = {}
        for s in candidates:
            key = (
                _length_to_mm(s.page_width, 210.0, min_mm=140.0, max_mm=360.0),
                _length_to_mm(s.page_height, 297.0, min_mm=200.0, max_mm=450.0),
                _length_to_mm(s.top_margin, 30.0, min_mm=8.0, max_mm=65.0),
                _length_to_mm(s.bottom_margin, 25.0, min_mm=8.0, max_mm=65.0),
                _length_to_mm(s.left_margin, 26.0, min_mm=8.0, max_mm=65.0),
                _length_to_mm(s.right_margin, 25.0, min_mm=8.0, max_mm=65.0),
            )
            geom_counts[key] += 1
            sec_by_geom[key] = s
        best_geom = geom_counts.most_common(1)[0][0]
        ref_sec = sec_by_geom[best_geom]
    else:
        ref_sec = doc.sections[0]

    fmt["page"]["width_mm"] = _length_to_mm(ref_sec.page_width, 210.0, min_mm=140.0, max_mm=360.0)
    fmt["page"]["height_mm"] = _length_to_mm(ref_sec.page_height, 297.0, min_mm=200.0, max_mm=450.0)
    fmt["page"]["margins_mm"] = {
        "top": _length_to_mm(ref_sec.top_margin, 30.0, min_mm=8.0, max_mm=65.0),
        "bottom": _length_to_mm(ref_sec.bottom_margin, 25.0, min_mm=8.0, max_mm=65.0),
        "left": _length_to_mm(ref_sec.left_margin, 26.0, min_mm=8.0, max_mm=65.0),
        "right": _length_to_mm(ref_sec.right_margin, 25.0, min_mm=8.0, max_mm=65.0),
    }

    title_fonts: Counter[tuple[str, float]] = Counter()
    h1_fonts: Counter[tuple[str, float]] = Counter()
    h2_fonts: Counter[tuple[str, float]] = Counter()
    h3_fonts: Counter[tuple[str, float]] = Counter()
    h4_fonts: Counter[tuple[str, float]] = Counter()
    body_fonts: Counter[tuple[str, float]] = Counter()
    line_spacings: Counter[float] = Counter()

    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if not txt or len(txt) < 2:
            continue
        fn, sz, _ = _paragraph_font(p)
        if not fn or not sz or not _is_cjk_capable_font(fn):
            continue
        jc = _paragraph_jc(p)
        if jc == "center" and sz >= 20.0 and len(txt) <= 45 and not txt.startswith("目"):
            title_fonts[(fn, sz)] += 1
        elif re.match(rf"^{_CN_NUM}[、.．]", txt):
            h1_fonts[(fn, sz)] += 1
        elif re.match(rf"^[（(]{_CN_NUM}[）)]", txt):
            h2_fonts[(fn, sz)] += 1
        elif re.match(r"^\d{1,2}[.．、]", txt) and len(txt) < 60:
            h3_fonts[(fn, sz)] += 1
        elif re.match(r"^[（(]\d{1,2}[）)]", txt) and len(txt) < 60:
            h4_fonts[(fn, sz)] += 1
        elif len(txt) >= 18:
            body_fonts[(fn, sz)] += 1
            ls = p.paragraph_format.line_spacing
            if ls is not None:
                pt = None
                if hasattr(ls, "pt"):
                    try:
                        pt = float(ls.pt)
                    except Exception:
                        pt = None
                elif isinstance(ls, (int, float)) and ls > 1000:
                    # Fallback if raw EMU value is passed directly
                    pt = float(ls) / 12700.0
                if pt is not None and 14.0 <= pt <= 44.0:
                    line_spacings[round(pt, 1)] += 1

    if title_fonts:
        fn, sz = title_fonts.most_common(1)[0][0]
        fmt["fonts"]["title"]["name"] = fn
        fmt["fonts"]["title"]["size_pt"] = sz
    if h1_fonts:
        fn, sz = h1_fonts.most_common(1)[0][0]
        fmt["fonts"]["heading1"]["name"] = fn
        fmt["fonts"]["heading1"]["size_pt"] = sz
    if h2_fonts:
        fn, sz = h2_fonts.most_common(1)[0][0]
        fmt["fonts"]["heading2"]["name"] = fn
        fmt["fonts"]["heading2"]["size_pt"] = sz
    if h3_fonts:
        fn, sz = h3_fonts.most_common(1)[0][0]
        fmt["fonts"].setdefault("heading3", {})["name"] = fn
        fmt["fonts"]["heading3"]["size_pt"] = sz
    if h4_fonts:
        fn, sz = h4_fonts.most_common(1)[0][0]
        fmt["fonts"].setdefault("heading4", {})["name"] = fn
        fmt["fonts"]["heading4"]["size_pt"] = sz
    if body_fonts:
        fn, sz = body_fonts.most_common(1)[0][0]
        fmt["fonts"]["body"]["name"] = fn
        fmt["fonts"]["body"]["size_pt"] = sz
    if line_spacings:
        fmt["fonts"]["body"]["line_spacing_pt"] = line_spacings.most_common(1)[0][0]

    tbl_fonts: Counter[tuple[str, float]] = Counter()
    for tbl in doc.tables[:20]:
        for row in tbl.rows[:5]:
            for cell in row.cells[:6]:
                for p in cell.paragraphs:
                    if p.text.strip():
                        fn, sz, _ = _paragraph_font(p)
                        if fn and sz and _is_cjk_capable_font(fn):
                            tbl_fonts[(fn, sz)] += 1
    if tbl_fonts:
        fn, sz = tbl_fonts.most_common(1)[0][0]
        fmt["fonts"]["table"]["name"] = fn
        fmt["fonts"]["table"]["size_pt"] = sz

    fmt["page_numbers"] = _detect_page_number_style(doc)
    return fmt


def _is_top_heading(p, sec_idx: int) -> bool:
    txt = re.sub(r"\s+", " ", (p.text or "").strip())
    if not txt or len(txt) > 45:
        return False
    style_name = (getattr(getattr(p, "style", None), "name", "") or "").lower()
    if "heading 1" in style_name or "标题 1" in style_name:
        return True
    p_pr = p._p.find(qn("w:pPr"))
    if p_pr is not None:
        ol = p_pr.find(qn("w:outlineLvl"))
        if ol is not None and ol.get(qn("w:val")) == "0":
            return True
    if re.match(rf"^{_CN_NUM}[、.．]\s*\S+", txt):
        return True
    if re.fullmatch(r"目\s*录|记\s*事|竞\s*赛\s*规\s*程|补\s*充\s*通\s*知|组\s*织\s*机\s*构", txt):
        return True
    # Standalone centered large title heading (e.g., 16pt+ short title not inside TOC)
    if sec_idx > 0 and _paragraph_jc(p) == "center" and len(txt) <= 24:
        _fn, sz, _bold = _paragraph_font(p)
        if sz and sz >= 16.0 and not re.search(r"20\d{2}年.*比赛|附件\d", txt):
            for _pat, _kind, _gen in _CHAPTER_KIND_RULES:
                if re.search(_pat, txt):
                    return True
    return False


def _p_text_with_tabs(p_el) -> str:
    parts: list[str] = []
    for node in p_el.iter():
        if node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag == qn("w:t") and node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def _clean_toc_line(raw: str) -> str:
    # Strip trailing tab + dot leaders + page numbers even if <w:tab> was followed directly by digits
    cleaned = re.sub(r"(\t|\.{2,}|…{1,})[\s.]*\d{1,4}\s*$", "", raw).strip()
    if cleaned == raw:
        cleaned = re.sub(r"(?<=[\u4e00-\u9fff）)])\s*\d{1,3}\s*$", "", raw).strip()
    return cleaned


def _extract_toc_candidates(doc: Document) -> list[str]:
    """Read TOC labels if the sample document already has a TOC block (used purely as chapter-outline names)."""
    entries: list[str] = []
    for sdt in doc._element.body.iter(qn("w:sdt")):
        for p_el in sdt.iter(qn("w:p")):
            txt = _clean_toc_line(_p_text_with_tabs(p_el))
            if txt and re.sub(r"\s+", "", txt) != "目录":
                entries.append(txt)
    if entries:
        return entries
    in_toc = False
    for p in doc.paragraphs[:150]:
        raw = _p_text_with_tabs(p._p)
        compact = re.sub(r"\s+", "", raw)
        if compact == "目录":
            in_toc = True
            continue
        if in_toc:
            if not raw:
                continue
            has_page_tail = bool(
                re.search(r"(\t|\.{2,}|…{1,})\s*\d{1,4}\s*$", raw)
                or re.search(r"[\u4e00-\u9fff）)]\d{1,3}\s*$", raw)
            )
            if has_page_tail:
                clean = _clean_toc_line(raw)
                if clean and re.sub(r"\s+", "", clean) != "目录":
                    entries.append(clean)
            elif len(entries) >= 2:
                break
    return entries


def _infer_roster_preset(doc: Document, chapters_meta: list[dict[str, Any]]) -> str:
    """Infer which of the 7 built-in roster_skill presets matches this sample order book."""
    roster_tables: list[tuple[list[str], int]] = []
    roster_paras: list[str] = []

    current_kind = ""
    sec_idx = 0
    for child in doc.element.body:
        if child.tag == qn("w:p"):
            txt = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            p_pr = child.find(qn("w:pPr"))
            if re.match(rf"^{_CN_NUM}[、.．]", txt) or any(
                re.search(pat, txt) for pat, k, _ in _CHAPTER_KIND_RULES if k == "roster"
            ):
                kind, _ = _classify_chapter_kind(txt)
                current_kind = kind
            elif txt and current_kind == "roster":
                roster_paras.append(txt)
            if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
                sec_idx += 1
        elif child.tag == qn("w:tbl") and current_kind == "roster":
            rows = child.findall(qn("w:tr"))
            if rows:
                first_cells = [
                    "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()
                    for tc in rows[0].findall(qn("w:tc"))
                ]
                roster_tables.append((first_cells, len(rows)))

    # Fallback: scan tables whose headers look like roster headers
    if not roster_tables and not roster_paras:
        for tbl in doc.tables:
            if not tbl.rows:
                continue
            first_cells = [c.text.strip() for c in tbl.rows[0].cells]
            joined = " ".join(first_cells)
            if any(
                k in joined
                for k in ("号码", "领队", "教练", "运动员", "参赛单位", "联系人", "联络员", "联系方式", "队员")
            ):
                roster_tables.append((first_cells, len(tbl.rows)))

    inline_hits = sum(
        1
        for txt in roster_paras[:40]
        if any(k in txt for k in ("领队：", "领队:", "教练：", "教练:", "队员：", "运动员："))
    )
    if inline_hits >= 2 and not roster_tables:
        return "inline-roster"

    for headers, _row_count in roster_tables:
        joined = " ".join(headers)
        norm_headers = [re.sub(r"\s+", "", h) for h in headers]
        if any(k in joined for k in ("联系人", "联系电话", "联系方式", "电话", "手机", "联络员")) and (
            "单位" in joined or "队名" in joined
        ) and "号码" not in joined:
            return "unit-contact-table"
        # Two-column numbered: e.g. 序号 姓名 号码 序号 姓名 号码 (4..8 cols)
        if len(headers) >= 4 and (
            norm_headers.count("姓名") >= 2 or norm_headers.count("号码") >= 2
        ):
            return "two-column-numbered"
        if ("单位" in joined or "代表队" in joined) and ("职务" in joined or "性别" in joined or "组别" in joined):
            if "领队" in " ".join(roster_paras[:15]) or "队员" in joined:
                return "team-wide"
        if "领队" in joined and ("队员" in joined or "运动员" in joined):
            return "grouped-team-table"

    if inline_hits >= 1:
        return "inline-roster"
    if roster_tables:
        return "team-wide"
    return "inline-roster"


_SUBARTICLE_RE = re.compile(
    r"主办单位|承办单位|协办单位|实施单位|比赛时间|比赛地点|活动时间|活动地点|活动内容|"
    r"参加单位|竞赛项目|项目规则|奖项设置|参赛要求|运动员资格|参加办法|竞赛办法|"
    r"录取名次|报名|纪律要求|裁判员|经费|未尽事宜|关于.*承诺|责任承担|"
    r"竞赛委员会|办事机构|仲裁委员会|赛前|赛中|赛后|文明观赛|最终解释权"
)


def _extract_structural_outline(doc: Document) -> list[dict[str, Any]]:
    """Extract clean, de-identified chapter outline from sample DOCX."""
    seen_kinds: list[tuple[str, str, bool]] = []  # (kind, title, landscape)

    sec_props = [
        (s.page_width or 0) > (s.page_height or 1) for s in doc.sections
    ]
    any_landscape = any(sec_props)
    sec_idx = 0
    has_cover_img = False
    has_back_img = False
    has_notes = False
    last_sec = len(doc.sections) - 1

    # Inspect TOC titles first as authoritative book-level chapter order
    toc_titles = _extract_toc_candidates(doc)
    has_toc = bool(toc_titles) or any(
        re.sub(r"\s+", "", (p.text or "").strip()) == "目录" for p in doc.paragraphs[:60]
    )

    body_headings: list[tuple[str, str, bool]] = []
    for child in doc.element.body:
        sec_landscape = sec_props[min(sec_idx, len(sec_props) - 1)] if sec_props else False
        if child.tag == qn("w:p"):
            txt = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            compact = re.sub(r"\s+", "", txt)
            has_img = (
                child.find(".//" + qn("a:blip")) is not None
                or child.find(".//{urn:schemas-microsoft-com:vml}imagedata") is not None
            )
            if sec_idx == 0 and has_img:
                has_cover_img = True
            if sec_idx == last_sec and sec_idx > 0 and has_img:
                has_back_img = True
            if compact == "记事":
                has_notes = True
            elif txt and len(txt) <= 45 and compact != "目录":
                if (
                    re.match(rf"^{_CN_NUM}[、.．]\s*\S+", txt)
                    or re.fullmatch(r"竞赛规程|组织机构|补充通知|分组情况|竞赛日程表|各代表队名单", compact)
                ) and not _SUBARTICLE_RE.search(compact):
                    kind, clean_title = _classify_chapter_kind(txt)
                    if kind not in {"toc", "notes"}:
                        body_headings.append((kind, clean_title, sec_landscape))
            p_pr = child.find(qn("w:pPr"))
            if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
                sec_idx += 1

    if len(toc_titles) >= 2:
        for title in toc_titles:
            kind, clean_title = _classify_chapter_kind(title)
            if kind in {"toc", "notes"}:
                if kind == "notes":
                    has_notes = True
                continue
            landscape = True if (kind == "schedule" and any_landscape) else False
            seen_kinds.append((kind, clean_title, landscape))
    else:
        seen_kinds = body_headings

    # Build deduplicated chapters list, collapsing group-split duplicates into split_by_group_in_toc=True
    chapters: list[dict[str, Any]] = [
        {
            "id": "cover",
            "title": "封面",
            "kind": "cover",
            "required": False,
            "page_break_before": False,
            "landscape": False,
            "split_by_group_in_toc": False,
            "toc_visible": False,
            "hint": "全幅封面图或赛事主画面底图",
        }
    ]
    if has_toc or seen_kinds:
        chapters.append(
            {
                "id": "toc",
                "title": "目录",
                "kind": "toc",
                "required": True,
                "page_break_before": True,
                "landscape": False,
                "split_by_group_in_toc": False,
                "toc_visible": False,
                "hint": "自动生成并同步正文页码",
            }
        )

    # Track counts of each kind to turn e.g. 2 roster entries into 1 split_by_group_in_toc roster slot
    counts_by_key: Counter[tuple[str, str]] = Counter((k, t) for k, t, _ in seen_kinds)
    added_keys: set[tuple[str, str]] = set()
    used_ids: set[str] = {"cover", "toc", "notes", "back_cover"}
    has_notes = False

    for idx, (kind, title, landscape) in enumerate(seen_kinds):
        if kind == "notes":
            has_notes = True
            continue
        key = (kind, title)
        if key in added_keys:
            continue
        added_keys.add(key)

        split_by_group = kind in {"roster", "standings", "brackets"} and counts_by_key[key] > 1
        cid_base = kind if kind not in used_ids else f"{kind}_{idx + 1}"
        while cid_base in used_ids:
            cid_base += "_x"
        used_ids.add(cid_base)

        chapters.append(
            {
                "id": cid_base,
                "title": title,
                "kind": kind,
                "required": kind in {"regulations", "roster"},
                "page_break_before": True,
                "landscape": "auto" if kind == "schedule" or landscape else False,
                "split_by_group_in_toc": split_by_group,
                "toc_visible": True,
                "hint": f"按样例模板“{title}”结构承接对应资料",
            }
        )

    # Guarantee essential fallback chapters if sample had minimal headings
    present_kinds = {c["kind"] for c in chapters}
    if "roster" not in present_kinds:
        chapters.append(
            {
                "id": "rosters",
                "title": "各代表队名单",
                "kind": "roster",
                "required": True,
                "page_break_before": True,
                "landscape": False,
                "split_by_group_in_toc": True,
                "toc_visible": True,
                "hint": "按参赛组别与报名表生成代表队名单",
            }
        )

    if has_notes or True:
        chapters.append(
            {
                "id": "notes",
                "title": "记事",
                "kind": "notes",
                "required": False,
                "page_break_before": True,
                "landscape": False,
                "split_by_group_in_toc": False,
                "toc_visible": False,
                "hint": "空白记事页",
            }
        )
    if has_back_img or True:
        chapters.append(
            {
                "id": "back_cover",
                "title": "尾页",
                "kind": "back_cover",
                "required": False,
                "page_break_before": True,
                "landscape": False,
                "split_by_group_in_toc": False,
                "toc_visible": False,
                "hint": "全幅封底图片",
            }
        )
    return chapters


def _refine_template_with_llm(spec: dict[str, Any]) -> dict[str, Any]:
    """Use LLM purely on stripped structural outline (never historical bodies) to refine labels/description."""
    api_key = getattr(config, "API_KEY", "")
    base_url = getattr(config, "BASE_URL", "").rstrip("/")
    model = getattr(config, "MODEL", getattr(config, "DEFAULT_MODEL", "qwen3.8-flash"))
    if not api_key or not base_url:
        return spec

    payload = {
        "model": model,
        "temperature": 0.1,
        "messages": [
            {
                "role": "system",
                "content": (
                    "你是体育赛事秩序册模板结构分析助手。你会收到一份已脱敏的模板结构草案"
                    "（仅含版式参数、roster_preset 与章节列表，不含具体赛事人员数据）。"
                    "请检查并优化每个非封面/目录/记事/尾页章节的 title，确保不包含具体年份、届数、地名、具体赛事名"
                    "（如若发现年份或具体比赛前缀则去掉，只保留通用章节名如'竞赛规程'、'补充通知'、'各代表队名单'等），"
                    "同时生成一段不超过80字的通用模板 description。"
                    "只返回合法 JSON: {\"description\": str, \"chapters\": [{\"id\": str, \"title\": str, \"kind\": str}]}"
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "name": spec["name"],
                        "roster_preset": spec["roster_preset"],
                        "chapters": [
                            {"id": c["id"], "title": c["title"], "kind": c["kind"]}
                            for c in spec["chapters"]
                        ],
                    },
                    ensure_ascii=False,
                ),
            },
        ],
    }
    req = urllib.request.Request(
        f"{base_url}/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=20.0) as resp:
            body = json.loads(resp.read().decode("utf-8"))
        content = body["choices"][0]["message"]["content"]
        match = re.search(r"\{.*\}", content, re.S)
        if not match:
            return spec
        refined = json.loads(match.group(0))
        if isinstance(refined.get("description"), str) and refined["description"].strip():
            spec["description"] = refined["description"].strip()
        allowed_kinds = {
            "cover",
            "toc",
            "richtext",
            "organization",
            "roster",
            "grouping",
            "schedule",
            "standings",
            "brackets",
            "notes",
            "back_cover",
            "custom",
        }
        by_id = {
            c.get("id"): c
            for c in (refined.get("chapters") or [])
            if isinstance(c, dict) and c.get("id")
        }
        for chap in spec["chapters"]:
            upd = by_id.get(chap["id"])
            if not upd:
                continue
            if chap["kind"] not in {"cover", "toc", "notes", "back_cover"}:
                new_title = str(upd.get("title") or "").strip()
                if new_title:
                    chap["title"] = new_title[:40]
                new_kind = str(upd.get("kind") or "").strip()
                if new_kind in allowed_kinds:
                    chap["kind"] = new_kind
    except Exception:
        return spec
    return spec


def parse_sample_book_to_template(
    sample_path: Path,
    *,
    name: str | None = None,
    use_llm: bool = True,
) -> dict[str, Any]:
    """Parse an uploaded Word sample book (.docx/.doc/.wps) into a stripped reusable TemplateSpec and save it."""
    sample_path = Path(sample_path).resolve()
    if not sample_path.is_file():
        raise FileNotFoundError(f"未找到样例文件: {sample_path}")

    with tempfile.TemporaryDirectory(prefix="tpl-parse-") as tmp_dir:
        tmp = Path(tmp_dir)
        docx_path = convert_to_docx(sample_path, tmp / "sample.docx")
        doc = Document(str(docx_path))

        formatting = _detect_formatting(doc)
        chapters = _extract_structural_outline(doc)
        roster_preset = _infer_roster_preset(doc, chapters)
        if roster_preset not in PRESETS:
            roster_preset = "inline-roster"

        watermark_path = _extract_watermark_asset(doc, tmp)

        raw_name = (name or "").strip() or sample_path.stem
        # Strip noisy trailing dates or "(1)" "(修订版)" for cleaner template name if user left blank
        clean_name = re.sub(r"[（(]\d+[）)]$", "", raw_name).strip() or "自定义秩序册模板"
        if not name:
            clean_name = f"模板：{clean_name}"

        chap_names = "、".join(c["title"] for c in chapters if c.get("toc_visible"))
        desc = (
            f"从样例《{sample_path.name}》提取：采用 {formatting['fonts']['body']['name']} "
            f"{formatting['fonts']['body']['size_pt']}pt、名单格式 `{roster_preset}`，"
            f"包含章节：{chap_names or '标准竞赛结构'}（已剥离原样例赛事历史数据）。"
        )

        spec: dict[str, Any] = {
            "name": clean_name,
            "source_filename": sample_path.name,
            "description": desc,
            "roster_preset": roster_preset,
            "formatting": formatting,
            "cover": {"mode": "image_or_bg"},
            "back_cover": {"mode": "image_or_skip"},
            "chapters": chapters,
        }
        if use_llm:
            spec = _refine_template_with_llm(spec)

        saved = save_custom_template(
            spec,
            watermark_src=watermark_path,
        )
        return saved
