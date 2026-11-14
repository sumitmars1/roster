from __future__ import annotations

import copy
import re
import shutil
import uuid
from pathlib import Path
from typing import Any

import config
from presets import PRESETS
from storage import now, read_json, write_json

TEMPLATES_DIR = config.DATA / "templates"


def _slugify(name: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9_-]+", "-", name).strip("-").lower()
    return slug or f"tpl-{uuid.uuid4().hex[:8]}"


def _default_official_formatting() -> dict[str, Any]:
    return {
        "page": {
            "width_mm": 210.0,
            "height_mm": 297.0,
            "margins_mm": {"top": 30.0, "bottom": 25.0, "left": 26.0, "right": 25.0},
            "landscape_margins_mm": {"top": 20.0, "bottom": 20.0, "left": 22.0, "right": 22.0},
            "grid": {"lines": 22, "chars": 28, "fill_box": True},
        },
        "fonts": {
            "title": {
                "name": "方正小标宋_GBK",
                "size_pt": 22.0,
                "bold": False,
                "align": "center",
                "top_blank_size_pt": 10.5,
                "bottom_blank_size_pt": 10.5,
            },
            "subtitle": {"name": "方正楷体_GBK", "size_pt": 16.0, "bold": False, "align": "center"},
            "heading1": {"name": "方正黑体_GBK", "size_pt": 16.0, "bold": False, "align": "left"},
            "heading2": {"name": "方正楷体_GBK", "size_pt": 16.0, "bold": False, "align": "left"},
            "heading3": {"name": "方正仿宋_GBK", "size_pt": 16.0, "bold": False, "align": "left"},
            "heading4": {"name": "方正仿宋_GBK", "size_pt": 16.0, "bold": False, "align": "left"},
            "attachment_label": {"name": "方正黑体_GBK", "size_pt": 16.0, "bold": False, "align": "left"},
            "body": {
                "name": "方正仿宋_GBK",
                "size_pt": 16.0,
                "bold": False,
                "line_spacing_pt": 30.0,
                "first_line_indent_chars": 2.0,
            },
            "table": {
                "name": "方正仿宋_GBK",
                "header_name": "方正黑体_GBK",
                "size_pt": 10.5,
                "header_size_pt": 10.5,
                "cell_pad_twips": 70,
            },
            "toc": {
                "title": "目    录",
                "title_name": "方正小标宋_GBK",
                "title_size_pt": 22.0,
                "entry_name": "方正仿宋_GBK",
                "entry_size_pt": 14.0,
                "line_spacing_pt": 24.0,
            },
        },
        "page_numbers": {
            "style": "hyphen",
            "prefix": "-",
            "suffix": "-",
            "font": "宋体",
            "size_pt": 14.0,
            "align": "odd_right_even_left",
            "indent_chars": 1.0,
            "odd_even": True,
            "skip_frontmatter": True,
            "skip_back_cover": True,
        },
        "toc_numbering": "arabic",
        "header_text": "",
        "watermark": {"enabled": False, "asset": None, "opacity": 0.18},
    }


def _chapter(
    cid: str,
    title: str,
    kind: str,
    *,
    required: bool = False,
    page_break_before: bool = True,
    landscape: bool | str = False,
    split_by_group_in_toc: bool = False,
    toc_visible: bool = True,
    hint: str = "",
) -> dict[str, Any]:
    return {
        "id": cid,
        "title": title,
        "kind": kind,
        "required": required,
        "page_break_before": page_break_before,
        "landscape": landscape,
        "split_by_group_in_toc": split_by_group_in_toc,
        "toc_visible": toc_visible,
        "hint": hint,
    }


def builtin_templates() -> list[dict[str, Any]]:
    fmt_std = _default_official_formatting()

    fmt_football = copy.deepcopy(fmt_std)
    fmt_football["fonts"]["table"]["size_pt"] = 10.0

    fmt_team_wide = copy.deepcopy(fmt_std)
    fmt_team_wide["toc_numbering"] = "chinese"

    fmt_fun_games = copy.deepcopy(fmt_team_wide)
    fmt_fun_games["fonts"]["toc"]["entry_size_pt"] = 14.0

    return [
        {
            "id": "official-inline",
            "name": "标准公文秩序册（横排代表队名单）",
            "builtin": True,
            "description": "严格遵循公文排版规范（A4、22行×28字版心、方正小标宋/黑体/楷体/仿宋、-1- 奇右偶左四号宋体页码），代表队名单采用横排紧凑结构，适用于篮球、象棋、拔河等常规单项或团体赛事。",
            "roster_preset": "inline-roster",
            "formatting": fmt_std,
            "chapters": [
                _chapter("cover", "封面", "cover", required=False, toc_visible=False, hint="全幅封面图或赛事主画面底图"),
                _chapter("toc", "目录", "toc", required=True, toc_visible=False, hint="自动生成并同步正文页码"),
                _chapter("regulations", "竞赛规程", "richtext", required=True, hint="竞赛规程及附件（含报名二维码、承诺书等）"),
                _chapter("supplement", "补充通知", "richtext", required=False, hint="规程补充说明或领队会纪要"),
                _chapter("organization", "组织机构", "organization", required=False, hint="组委会、办事机构、仲裁委员会、裁判组等"),
                _chapter(
                    "rosters",
                    "各代表队名单",
                    "roster",
                    required=True,
                    split_by_group_in_toc=True,
                    hint="按竞赛规程与分组确定的各组别代表队名单",
                ),
                _chapter("grouping", "分组情况", "grouping", required=False, hint="抽签分组表"),
                _chapter("schedule", "竞赛日程表", "schedule", required=False, landscape="auto", hint="竞赛日程/赛程表"),
                _chapter(
                    "standings",
                    "比赛积分表",
                    "standings",
                    required=False,
                    split_by_group_in_toc=True,
                    hint="小组循环积分表",
                ),
                _chapter(
                    "brackets",
                    "淘汰赛对阵表",
                    "brackets",
                    required=False,
                    split_by_group_in_toc=True,
                    hint="交叉淘汰及决赛对阵图/表",
                ),
                _chapter("guide", "参赛指南", "richtext", required=False, hint="报到须知、场地交通、安全提示"),
                _chapter("notes", "记事", "notes", required=False, toc_visible=False, hint="空白记事页"),
                _chapter("back_cover", "尾页", "back_cover", required=False, toc_visible=False, hint="全幅封底图片"),
            ],
            "created_at": "builtin",
            "updated_at": "builtin",
        },
        {
            "id": "official-football",
            "name": "球类联赛秩序册（双栏编号名单+宽表横排赛程）",
            "builtin": True,
            "description": "适用于五人制/十一人制足球及需要号码编号的球类赛事；名单采用双栏带号码表格，支持补充通知、横版宽表赛程与积分榜。",
            "roster_preset": "two-column-numbered",
            "formatting": fmt_football,
            "chapters": [
                _chapter("cover", "封面", "cover", toc_visible=False),
                _chapter("toc", "目录", "toc", required=True, toc_visible=False),
                _chapter("regulations", "竞赛规程", "richtext", required=True),
                _chapter("supplement", "补充通知", "richtext"),
                _chapter("organization", "组织机构", "organization"),
                _chapter("rosters", "各代表队名单", "roster", required=True, split_by_group_in_toc=True),
                _chapter("grouping", "分组情况", "grouping"),
                _chapter("schedule", "竞赛日程表", "schedule", landscape="auto"),
                _chapter("standings", "比赛积分表", "standings", split_by_group_in_toc=True),
                _chapter("brackets", "淘汰赛对阵表", "brackets", split_by_group_in_toc=True),
                _chapter("notes", "记事", "notes", toc_visible=False),
                _chapter("back_cover", "尾页", "back_cover", toc_visible=False),
            ],
            "created_at": "builtin",
            "updated_at": "builtin",
        },
        {
            "id": "team-wide-event",
            "name": "综合赛事/系统运动会秩序册（团体宽表+指南安排）",
            "builtin": True,
            "description": "适用于系统运动会、企事业单位联谊赛等综合性赛事，名单使用分队表格（含职务/号码/备注），包含日程安排、场地示意图与安全应急预案。",
            "roster_preset": "team-wide",
            "formatting": fmt_team_wide,
            "chapters": [
                _chapter("cover", "封面", "cover", toc_visible=False),
                _chapter("toc", "目录", "toc", required=True, toc_visible=False),
                _chapter("guide", "活动日程与参赛指南", "richtext", hint="开幕日程、报到乘车、场地交通等指南"),
                _chapter("regulations", "竞赛规程", "richtext", required=True),
                _chapter("organization", "组织机构及裁判名单", "organization"),
                _chapter("rosters", "各代表队名单", "roster", required=True, split_by_group_in_toc=True),
                _chapter("grouping", "分组情况", "grouping"),
                _chapter("schedule", "竞赛日程表", "schedule", landscape="auto"),
                _chapter("standings", "比赛积分表", "standings", split_by_group_in_toc=True),
                _chapter("brackets", "淘汰赛对阵表", "brackets", split_by_group_in_toc=True),
                _chapter("safety", "场地分布与应急预案", "richtext", hint="平面图、联系人及医疗应急预案"),
                _chapter("notes", "记事", "notes", toc_visible=False),
                _chapter("back_cover", "尾页", "back_cover", toc_visible=False),
            ],
            "created_at": "builtin",
            "updated_at": "builtin",
        },
        {
            "id": "fun-games-contacts",
            "name": "趣味运动会秩序册（单位联络表+轮转编排）",
            "builtin": True,
            "description": "适用于趣味运动会、开幕式暨嘉年华类赛事，支持总体方案、整体流程、参赛单位联络名单表、轮转项目编排、站位/平面图与安全须知。",
            "roster_preset": "unit-contact-table",
            "formatting": fmt_fun_games,
            "chapters": [
                _chapter("cover", "封面", "cover", toc_visible=False),
                _chapter("toc", "目录", "toc", required=True, toc_visible=False),
                _chapter("regulations", "总体方案", "richtext", required=True, hint="总体方案及竞赛项目说明"),
                _chapter("guide", "整体流程", "richtext", hint="开幕式及活动整体时间流程"),
                _chapter("organization", "组织机构", "organization"),
                _chapter("rosters", "参赛单位名单", "roster", required=True),
                _chapter("schedule", "竞赛安排", "schedule", landscape="auto", hint="项目轮转与场次安排"),
                _chapter("venue", "站位及平面分布图", "richtext", hint="队伍站位图、场地平面分布图"),
                _chapter("safety", "注意事项及安全须知", "richtext", hint="安全提示、医疗应急、服务保障"),
                _chapter("notes", "记事", "notes", toc_visible=False),
                _chapter("back_cover", "尾页", "back_cover", toc_visible=False),
            ],
            "created_at": "builtin",
            "updated_at": "builtin",
        },
    ]


def ensure_builtin_templates() -> None:
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)


def list_templates() -> list[dict[str, Any]]:
    ensure_builtin_templates()
    builtins = {t["id"]: t for t in builtin_templates()}
    results: list[dict[str, Any]] = []
    for item in builtin_templates():
        summary = copy.deepcopy(item)
        summary["asset_dir"] = None
        results.append(summary)

    for folder in sorted(TEMPLATES_DIR.iterdir()):
        if not folder.is_dir():
            continue
        meta_path = folder / "template.json"
        if not meta_path.is_file():
            continue
        data = read_json(meta_path)
        tid = data.get("id") or folder.name
        if tid in builtins:
            continue
        data["builtin"] = False
        results.append(data)
    return results


def get_template(template_id: str) -> dict[str, Any] | None:
    for item in builtin_templates():
        if item["id"] == template_id:
            return copy.deepcopy(item)
    meta_path = TEMPLATES_DIR / template_id / "template.json"
    if meta_path.is_file():
        data = read_json(meta_path)
        data["builtin"] = False
        return data
    return None


def template_asset_path(template_id: str, rel_name: str | None) -> Path | None:
    if not rel_name:
        return None
    candidate = TEMPLATES_DIR / template_id / "assets" / rel_name
    if candidate.is_file():
        return candidate
    return None


def save_custom_template(
    spec: dict[str, Any],
    *,
    watermark_src: Path | None = None,
    cover_bg_src: Path | None = None,
    sample_docx_src: Path | None = None,
) -> dict[str, Any]:
    ensure_builtin_templates()
    builtins = {t["id"] for t in builtin_templates()}
    raw_id = str(spec.get("id") or "").strip()
    if not raw_id or raw_id in builtins:
        raw_id = f"custom-{_slugify(str(spec.get('name') or 'template'))}-{uuid.uuid4().hex[:6]}"
    else:
        raw_id = _slugify(raw_id)
        if raw_id in builtins:
            raw_id = f"{raw_id}-{uuid.uuid4().hex[:6]}"

    t_dir = TEMPLATES_DIR / raw_id
    assets_dir = t_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    cleaned = copy.deepcopy(spec)
    cleaned["id"] = raw_id
    cleaned["builtin"] = False
    roster_preset = str(cleaned.get("roster_preset") or "inline-roster")
    if roster_preset not in PRESETS:
        roster_preset = "inline-roster"
    cleaned["roster_preset"] = roster_preset

    fmt = cleaned.setdefault("formatting", _default_official_formatting())
    wm = fmt.setdefault("watermark", {"enabled": False, "asset": None, "opacity": 0.18})

    if watermark_src and Path(watermark_src).is_file():
        wm_name = f"watermark{Path(watermark_src).suffix.lower() or '.png'}"
        shutil.copy2(watermark_src, assets_dir / wm_name)
        wm["enabled"] = True
        wm["asset"] = wm_name

    if cover_bg_src and Path(cover_bg_src).is_file():
        bg_name = f"cover_bg{Path(cover_bg_src).suffix.lower() or '.png'}"
        shutil.copy2(cover_bg_src, assets_dir / bg_name)
        cleaned.setdefault("cover", {})["default_bg_asset"] = bg_name

    if sample_docx_src and Path(sample_docx_src).is_file():
        shutil.copy2(sample_docx_src, t_dir / "sample_original.docx")

    stamp = now()
    cleaned.setdefault("created_at", stamp)
    cleaned["updated_at"] = stamp
    write_json(t_dir / "template.json", cleaned)
    return cleaned


def delete_custom_template(template_id: str) -> bool:
    for item in builtin_templates():
        if item["id"] == template_id:
            return False
    t_dir = TEMPLATES_DIR / template_id
    if not t_dir.is_dir():
        return False
    shutil.rmtree(t_dir)
    return True
