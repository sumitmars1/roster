"""Fill built-in roster layouts that take section-based input."""

from copy import deepcopy
import re
import unicodedata
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from roster import (
    BODY, W, XML_SPACE, array, load_json, parse_xml, required_text,
    save_json, separate_paths, temporary_output,
)
from roster_table_layout import adapt_regular_tables, child, prepare_team_table, usable_width


ASSETS = Path(__file__).resolve().parents[1] / "assets"
SLOT = re.compile(r"\{\{([a-z][a-z0-9_]*)(\?)?\}\}")


# Saved inputs may refer to the former sample names and fixed-width variants.
INLINE_ALIASES = {
    "three-across", "six-across-contacts", "four-across-numbered", "six-across-coach",
    "qijiang-chess", "qijiang-tug", "liangjiang-football", "basketball-inline",
}
SECTION_ALIASES = {
    **dict.fromkeys(INLINE_ALIASES, "inline-roster"),
    "happy-valley-basketball": "grouped-team-table",
    "municipal-fun": "unit-contact-table",
    "system-basketball": "team-wide",
    "team-wide-contacts": "team-wide",
}
SELECTION_ALIASES = {
    **dict.fromkeys(("tug", "long-rope", "football", "basketball"), "team-wide"),
    **dict.fromkeys(("badminton-singles", "table-tennis-singles"), "individual-table"),
    **dict.fromkeys(("badminton-doubles", "table-tennis-doubles"), "paired-table"),
    **dict.fromkeys(("badminton-contacts", "table-tennis-contacts"), "unit-contact-table"),
}
LEGACY_LAYOUTS = {
    "team-wide": {"men", "women", "12-slots", "10-slots-two-rows",
                  "13-slots-one-row", "9-slots-two-rows"},
    "individual-table": {"wide-sign"},
    "paired-table": {"wide-sign"},
    "unit-contact-table": {"landscape", "landscape-wide-contact"},
}
PLAYER_HEADINGS = {"men": "男运动员", "women": "女运动员"}
REGULAR_TABLES = {"individual-table", "paired-table", "unit-contact-table", "grouped-team-table"}
# Person-name cells follow the sample: a two-character name takes the width of
# three characters, so two- and three-character names line up on both sides.
NAME_SLOTS = {"name", "partner", "contact", "leader"}
PLAYER_SLOT = re.compile(r"player\d+(?:_numbered)?")
STAFF_FIELDS = {
    "leader": "领队", "liaison": "联络员", "coach": "教练", "coach_2": "教练",
    "assistant_1": "助理教练", "assistant_2": "助理教练", "assistant_3": "助理教练",
    "medic": "队医",
}


def section_presets():
    return sorted(folder.name for folder in ASSETS.iterdir()
                  if folder.is_dir() and any(folder.glob("*.docx")))


class Record:
    def __init__(self, data, path, source="", sequence=1, parent=None):
        self.data = data
        self.path = path
        self.source = required_text(data["source"], f"{path}.source") if "source" in data else source
        if "needs_review" in data and data["needs_review"] is not False:
            raise ValueError(f"{path} 的资料需复核（{self.source}）")
        self.sequence = sequence
        self.parent = parent
        self.page_width = parent.page_width if parent else None
        self.used = {"source", "needs_review"}

    def get(self, key, default=None):
        self.used.add(key)
        if key == "seq":
            return str(self.sequence)
        return self.data.get(key, default)

    def verify(self):
        remaining = sorted(key for key in self.data.keys() - self.used
                           if key == "seq" or self.data[key] not in (None, "", []))
        if remaining:
            raise ValueError(f"{self.path} 的字段在此母版没有位置：{', '.join(remaining)}")


def render_control(node, record, audit):
    command, key, *options = node.find(f"{W}sdtPr/{W}tag").get(W + "val").split(":")
    pattern = list(node.find(W + "sdtContent"))
    if command == "optional":
        # The nested slot consumes the value when present.
        if record.get(key) in (None, ""):
            return []
        return render_nodes(pattern, record, audit)

    rows = array(record.get(key, []), f"{record.path}.{key}")
    if not rows and not command.endswith("-optional"):
        raise ValueError(f"{record.path}.{key} 没有记录")
    if command == "matrix":
        width = int(options[0])
        return render_matrix(pattern, record, key, rows, width, audit)
    if command == "flow":
        return render_inline_players(pattern, record, key, rows, audit)

    output = []
    for index, data in enumerate(rows, 1):
        child = Record(data, f"{record.path}.{key}[{index}]", record.source, index, parent=record)
        nodes = render_nodes(pattern if command != "alternate" else [pattern[min(index - 1, 1)]], child, audit)
        if command != "project":
            child.verify()
        if command == "repeat-page" and index > 1:
            properties = nodes[0].find(W + "pPr")
            flag = properties.find(W + "pageBreakBefore")
            if flag is None:
                flag = etree.SubElement(properties, W + "pageBreakBefore")
            flag.set(W + "val", "1")
        output.extend(nodes)
    return output


SPACES = re.compile(r"[\s\u3000]+")


def inline_name(name):
    # 原表常用空格对齐两字名，抽取时会带进姓名（如“戴  池”）；先压掉空格再按两字补全角空格，
    # 这样每个姓名都占满三字宽，两字名与三字名左右都能对齐。
    compact = SPACES.sub("", name)
    return "\u3000".join(compact) if han_name(compact, 2) else compact


def set_player_tabs(paragraph, stops):
    ppr = paragraph.find(W + "pPr")
    ind = ppr.find(W + "ind")
    if ind is not None:
        ind.set(W + "left", "0")  # 各行都从页边起算，列位交给制表位，不再依赖行首缩进
    tabs = ppr.find(W + "tabs")
    if tabs is None:
        tabs = etree.Element(W + "tabs")
        snap = ppr.find(W + "snapToGrid")
        ppr.insert(ppr.index(snap) if snap is not None else len(ppr), tabs)
    for position in stops:
        stop = etree.SubElement(tabs, W + "tab")
        stop.set(W + "val", "left")
        stop.set(W + "pos", str(position))


def render_inline_players(pattern, parent, key, rows, audit):
    marker = "{{player_line?}}"
    first, following = pattern
    name_run = next(t.getparent() for t in first.iter(W + "t") if t.text == marker)
    em = 10 * int(name_run.find(f"{W}rPr/{W}sz").get(W + "val"))
    left = int(following.find(f"{W}pPr/{W}ind").get(W + "left"))
    available = parent.page_width - em // 4
    column = 7 * em  # 每列固定 7 字宽，够放“三字名 + 两位号码”
    columns = max(1, min(3, available // column))
    lines = []
    for index, data in enumerate(rows, 1):
        player = Record(data, f"{parent.path}.{key}[{index}]", parent.source)
        name = required_text(player.get("name"), f"{player.path}.name")
        number = player.get("number")
        if number not in (None, ""):
            number = required_text(number, f"{player.path}.number")
        player.verify()
        displayed = inline_name(name)
        if number:
            displayed += f"（{number}）"
        if not lines or len(lines[-1]) >= columns:
            lines.append([])
        lines[-1].append(displayed)
        audit.append({"record": player.path, "field": "name", "value": name, "source": player.source,
                      "row": len(lines), "column": len(lines[-1]),
                      **({"number": number} if number else {})})
    stops = [left + column * step for step in range(columns)]
    output = []
    for index, names in enumerate(lines):
        paragraph = deepcopy(following if index else first)
        field = next(t for t in paragraph.iter(W + "t") if t.text == marker)
        run = field.getparent()
        run.remove(field)
        if index:
            etree.SubElement(run, W + "tab")  # 后续行先跳到首列
        for offset, text in enumerate(names):
            if offset:
                etree.SubElement(run, W + "tab")
            item = etree.SubElement(run, W + "t")
            item.text = text
            item.set(XML_SPACE, "preserve")
        set_player_tabs(paragraph, stops)
        output.append(paragraph)
    return output


def render_matrix(pattern, parent, key, rows, width, audit):
    output = []
    inline = pattern[0].tag == W + "p"
    numbered = any("_numbered" in (node.text or "") for prototype in pattern for node in prototype.iter(W + "t"))
    for start in range(0, len(rows), width):
        fields = {}
        for column, data in enumerate(rows[start:start + width], 1):
            player = Record(data, f"{parent.path}.{key}[{start + column}]", parent.source)
            name = required_text(player.get("name"), f"{player.path}.name")
            number = player.get("number") if numbered else None
            if number not in (None, ""):
                number = required_text(number, f"{player.path}.number")
            player.verify()
            field = f"player{column}"
            displayed = "  ".join(name) if inline and han_name(name, 2) else name
            fields[field + "_numbered" if numbered else field] = (
                f"{displayed}（{number}）" if number else displayed
            )
            audit.append({"record": player.path, "field": "name", "value": name, "source": player.source,
                          "row": start // width + 1, "column": column, **({"number": number} if number else {})})
        prototype = pattern[min(start // width, len(pattern) - 1)]
        output.extend(render_nodes([prototype], parent, audit, matrix_fields=fields))
        if fields:
            raise ValueError(f"{parent.path}.{key} 的母版缺少姓名格位")
    return output


def render_nodes(prototypes, record, audit, *, matrix_fields=None):
    output = []
    for prototype in prototypes:
        if prototype.tag == W + "sdt":
            output.extend(render_control(prototype, record, audit))
            continue
        node = deepcopy(prototype)
        namespace = {"w": W[1:-1]}
        for field in node.xpath(".//w:t[not(ancestor::w:sdt)]", namespaces=namespace):
            match = SLOT.fullmatch(field.text or "")
            if not match:
                continue
            key, optional = match.groups()
            owner = record
            if key == "parent_group":
                owner, key = record.parent, "group"
            spread = key in ("staff_role", "staff_name")
            if spread:
                key = key.removeprefix("staff_")
            is_matrix = matrix_fields is not None and re.fullmatch(r"player\d+(?:_numbered)?", key)
            raw = matrix_fields.pop(key, None) if is_matrix else owner.get(key)
            value = "" if optional and raw in (None, "") else required_text(raw, f"{owner.path}.{key}")
            field.text = inline_name(value) if spread else value
            field.set(XML_SPACE, "preserve")
            if value and not spread and (key in NAME_SLOTS or PLAYER_SLOT.fullmatch(key)):
                fit_name(field, value)
            if value and not is_matrix:
                audit.append({"record": owner.path, "field": key, "value": value, "source": owner.source})
        for control in node.xpath(".//w:sdt[not(ancestor::w:sdt)]", namespaces=namespace):
            parent = control.getparent()
            index = parent.index(control)
            parent.remove(control)
            for offset, replacement in enumerate(render_control(control, record, audit)):
                parent.insert(index + offset, replacement)
        output.append(node)
    return output


def han_name(name, length):
    return bool(name) and len(name) == length and all(
        unicodedata.name(character, "").startswith("CJK ") for character in name
    )


def pad_name(value):
    """Give a two-character name the width of three characters."""
    return "\u3000".join(value) if han_name(value, 2) else value


def fit_name(field, value):
    """Match the sample, where a two-character name occupies three characters' width."""
    if not han_name(value, 2):
        return
    paragraph = next(field.iterancestors(W + "p"), None)
    if paragraph is not None:
        field.text = pad_name(value)


def format_wide_table_names(table, *, contacts, has_title):
    header_rows = 2 if has_title else 1
    for row in table.findall(W + "tr")[header_rows:]:
        cells = row.findall(W + "tc")
        for cell in (cells[2:3] + cells[4:] if contacts else cells[2:]):
            paragraph = cell.find(W + "p")
            run = paragraph.find(W + "r")
            value = run.find(W + "t") if run is not None else None
            name = value.text if value is not None else ""
            if not han_name(name, 4):
                continue
            # Keep the first two and last two characters together, as in the sample.
            following = deepcopy(paragraph)
            following.find(f"{W}r/{W}t").text = pad_name(name[2:])
            value.text = pad_name(name[:2])
            cell.append(following)


def normalize_inline_staff(staff):
    def lines(value):
        if not isinstance(value, str):
            return []
        return [line.strip() for line in re.split(r"\r\n|[\r\n\t]", value)]

    for entry in staff:
        names = lines(entry.get("name"))
        phones = lines(entry.get("phone"))
        if len(names) <= 1 and len(phones) <= 1:
            yield entry
        elif len(names) > 1 and len(names) == len(phones) and all(names) and all(phones):
            for name, phone in zip(names, phones):
                yield {**entry, "name": name, "phone": phone}
        else:
            # Keep unpaired values together instead of guessing which person owns a phone.
            normalized = dict(entry)
            if len(names) > 1:
                normalized["name"] = "、".join(filter(None, names))
            if len(phones) > 1:
                normalized["phone"] = "、".join(filter(None, phones))
            yield normalized


def normalize_inline(fields, label, *, directory=False):
    show_directory = fields.pop("show_directory", directory)
    if not isinstance(show_directory, bool):
        raise ValueError(f"{label}.show_directory 须为布尔值")
    if "teams" in fields:
        group = fields.pop("title", "")
        fields["groups"] = [{"group": group, "teams": fields.pop("teams")}]

    groups = []
    for group in array(fields.get("groups"), f"{label}.groups"):
        group = dict(group)
        if "title" in group and "group" not in group:
            group["group"] = group.pop("title")
        teams = []
        for original in array(group.get("teams"), f"{label}.teams"):
            team = dict(original)
            if "name" in team and "team" not in team:
                team["team"] = team.pop("name")
            if isinstance(team.get("team"), str):
                # A line break in a unit cell is a display separator, not a second team record.
                team["team"] = re.sub(r"[ \t]*[\r\n\t]+[ \t]*", " ", team["team"]).strip()
            staff = list(array(team.pop("staff", []), f"{label}.staff"))
            sources = team.pop("field_sources", {})
            for field, role in STAFF_FIELDS.items():
                name = team.pop(field, None)
                phone = team.pop(f"{field}_phone", None)
                if name or phone:
                    entry = {"role": role, "name": name, "phone": phone}
                    if sources.get(field):
                        entry["source"] = sources[field]
                    staff.append(entry)
            team["staff"] = list(normalize_inline_staff(staff))
            teams.append(team)
        group["teams"] = teams
        if show_directory:
            group["directory"] = [{"team": team["team"], **({"source": team["source"]} if team.get("source") else {})}
                                  for team in teams]
        groups.append(group)
    fields["groups"] = groups
    return fields


def section_template(preset, layout, data, label):
    directory = SECTION_ALIASES.get(preset, preset)
    if preset == "metro-selection":
        if layout not in SELECTION_ALIASES:
            raise ValueError(f"{label} 的 layout 须为 {', '.join(SELECTION_ALIASES)}")
        directory = SELECTION_ALIASES[layout]
    elif layout != "default" and layout not in LEGACY_LAYOUTS.get(directory, ()):
        raise ValueError(f"{label} 使用未识别的 layout：{layout}")

    if directory not in section_presets():
        raise ValueError(f"未知内置母版 {preset!r}，可选 {', '.join(section_presets())}")
    fields = data.copy()
    if directory == "inline-roster":
        normalize_inline(fields, label, directory=preset in ("four-across-numbered", "liangjiang-football"))
    if directory == "team-wide":
        fields.setdefault("player_heading", PLAYER_HEADINGS.get(layout, "运动员"))
    if directory == "unit-contact-table":
        heading = "趣味运动会参赛单位名单" if preset == "municipal-fun" else "参赛单位名单"
        fields.setdefault("title", heading)
        if layout.startswith("landscape") or preset == "metro-selection":
            fields.setdefault("orientation", "landscape")
        if "rows" in fields:
            rows = array(fields["rows"], f"{label}.rows")
            contacts = []
            for row in rows:
                if "leader" in row and "contact" not in row:
                    row = row.copy()
                    row["contact"] = row.pop("leader")
                contacts.append(row)
            fields["rows"] = contacts

    return ASSETS / directory / "default.docx", fields


def read_sections(preset, paths):
    sections = []
    for path in paths:
        payload = load_json(path)
        if set(payload) == {"groups"} and (preset == "inline-roster" or preset in INLINE_ALIASES):
            asset, fields = section_template(preset, "default", payload, path.name)
            sections.append((asset, Record(fields, path.name, path.name)))
            continue
        if set(payload) != {"sections"}:
            raise ValueError(f"{path.name} 须提供 sections 数组")
        for index, entry in enumerate(array(payload["sections"], f"{path.name}.sections"), 1):
            label = f"{path.name}.sections[{index}]"
            layout = entry.get("layout", "default")
            if not isinstance(layout, str):
                raise ValueError(f"{label} 的 layout 须为文本")
            if set(entry) - {"layout", "data"} or not isinstance(entry.get("data"), dict):
                raise ValueError(f"{label} 须提供 data 对象与可选 layout")
            asset, fields = section_template(preset, layout, entry["data"], label)
            sections.append((asset, Record(fields, label, path.name)))
    if not sections:
        raise ValueError("没有名单分段")
    return sections


def render_sections(args, *, groups=None):
    if groups is None:
        sections = read_sections(args.preset, args.data)
    else:
        asset, fields = section_template(args.preset, "default", {"groups": groups}, "名单")
        sections = [(asset, Record(fields, "名单"))]
    separate_paths([*(args.data or []), *ASSETS.rglob("*.docx")], [args.output, args.audit])
    audit = {"preset": args.preset, "sections": []}
    with ZipFile(sections[0][0]) as base:
        document = parse_xml(base.read(BODY))
        body = document.find(W + "body")
        for node in list(body):
            body.remove(node)
        for index, (asset, record) in enumerate(sections):
            with ZipFile(asset) as template:
                prototype = parse_xml(template.read(BODY)).find(W + "body")
            *content, section = list(prototype)
            record.page_width = usable_width(section)
            bindings = []
            layout = asset.parent.name
            widths = None
            if layout == "team-wide":
                teams = array(record.get("teams"), f"{record.path}.teams")
                if not teams:
                    raise ValueError(f"{record.path}.teams 没有队伍")
                for team in teams:
                    array(team.get("players", []), f"{record.path}.players")
                table = prototype.find(W + "tbl")
                contacts = any(team.get("phone") for team in teams)
                has_title = bool(record.get("title"))
                widths = prepare_team_table(table, section, teams, contacts, has_title)
            rendered = render_nodes(content, record, bindings)
            if layout == "team-wide":
                for table in (table for node in rendered for table in node.iter(W + "tbl")):
                    format_wide_table_names(table, contacts=contacts, has_title=has_title)
            if layout in REGULAR_TABLES:
                orientation = None
                if layout == "unit-contact-table":
                    orientation = record.get("orientation", "auto")
                    if orientation not in (None, "auto", "portrait", "landscape"):
                        raise ValueError(f"{record.path}.orientation 须为 auto、portrait 或 landscape")
                adapt_regular_tables(rendered, section, orientation)
            body.extend(rendered)
            record.verify()
            entry = {"layout": asset.stem, "bindings": bindings}
            if widths is not None:
                entry["column_widths_twips"] = widths
            audit["sections"].append(entry)
            if index == len(sections) - 1:
                body.append(deepcopy(section))
            else:
                paragraph = etree.SubElement(body, W + "p")
                properties = etree.SubElement(paragraph, W + "pPr")
                properties.append(deepcopy(section))
        xml = etree.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
        with temporary_output(args.output) as docx, temporary_output(args.audit) as report:
            with ZipFile(docx, "w") as output:
                for part in base.infolist():
                    output.writestr(part, xml if part.filename == BODY else base.read(part))
            save_json(report, audit)
            docx.replace(args.output)
            report.replace(args.audit)
    print(f"已生成 {len(sections)} 个名单分段")
