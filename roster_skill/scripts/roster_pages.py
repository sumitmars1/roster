"""Generate standalone roster pages from the built-in DOCX catalog."""

from copy import deepcopy
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree
from openpyxl import load_workbook

from roster import BODY, W, array, load_json, parse_xml, required_text, save_json, separate_paths, temporary_output
from roster_sections import ASSETS, INLINE_ALIASES, render_sections, section_presets


TEAM_PRESETS = ("two-column-numbered",)
TEAM_ALIASES = {"football-table": "two-column-numbered"}
MARKER = re.compile(r"\{\{([a-z_0-9]+)\}\}")
STAFF = ("leader", "coach", "coach_2", "assistant_1", "assistant_2", "assistant_3", "medic")
DETAILS = (*STAFF, "leader_phone", "coach_phone", "affiliation")
COLUMN_NAMES = {
    "leader": ("领队",),
    "coach": ("主教练", "教练"),
    "leader_phone": ("领队电话",),
    "coach_phone": ("教练电话", "主教练电话"),
    "assistant_1": ("助理教练1", "助理教练"),
    "assistant_2": ("助理教练2",),
    "assistant_3": ("助理教练3",),
    "medic": ("队医",),
}


def add_command(commands):
    cmd = commands.add_parser("pages", help="用内置母版生成独立名单页")
    presets = (*TEAM_PRESETS, *section_presets())
    cmd.add_argument("--preset", required=True, metavar="{" + ",".join(presets) + "}")
    source = cmd.add_mutually_exclusive_group(required=True)
    source.add_argument("--xlsx", type=Path)
    source.add_argument("--data", type=Path, action="append", help="按母版字段整理的名单 JSON")
    source.add_argument("--manifest", type=Path, help="按清单顺序合并 Excel 和已提取的名单 JSON")
    cmd.add_argument("--sheet", action="append", default=[], metavar="工作表=组别")
    cmd.add_argument("--output", type=Path, required=True)
    cmd.add_argument("--audit", type=Path, required=True)
    cmd.set_defaults(run=render_pages)


def source_text(value):
    return "" if value is None else str(value).strip()


def load_excel_groups(path, selection, preset):
    if not selection:
        raise ValueError("请用 --sheet 工作表=组别 指定要输出的工作表")
    book = load_workbook(path, data_only=True, read_only=True)
    try:
        groups = []
        used = set()
        for selector in selection:
            sheet, separator, group = selector.partition("=")
            if not separator or not group.strip() or sheet not in book.sheetnames:
                raise ValueError(f"无效的 --sheet {selector!r}，请使用 工作表=组别")
            if sheet in used:
                raise ValueError(f"重复选择工作表 {sheet}")
            used.add(sheet)
            teams = read_sheet(book[sheet], f"{path.name} / {sheet}", preset)
            groups.append({"title": group, "teams": teams})
        return groups
    finally:
        book.close()


def read_sheet(sheet, origin, preset):
    rows = sheet.iter_rows(values_only=True)
    headers = [source_text(value) for value in next(rows, ())]
    if headers.count("单位") != 1 or headers.count("姓名") != 1:
        raise ValueError(f"{origin} 第1行需要各有一列「单位」和「姓名」")
    team_column, name_column = headers.index("单位"), headers.index("姓名")
    number_column = next((i for i, key in enumerate(headers) if key in ("球衣", "球衣号码", "号码")), None)
    staff_columns = {
        field: next((headers.index(key) for key in keys if key in headers), None)
        for field, keys in COLUMN_NAMES.items()
    }
    if preset == "inline-roster":
        for field, heading in (("liaison", "联络员"), ("liaison_phone", "联络员电话"), ("medic_phone", "队医电话")):
            staff_columns[field] = headers.index(heading) if heading in headers else None
    generic_phones = [i for i, label in enumerate(headers) if label == "电话"]
    for field, index in zip(("leader_phone", "coach_phone"), generic_phones):
        if staff_columns[field] is None:
            staff_columns[field] = index

    teams = []
    team = None
    for index, row in enumerate(rows, 2):
        reference = f"{origin} / 第{index}行"
        unit = source_text(row[team_column])
        name = source_text(row[name_column])
        if unit:
            team = {"name": "、".join(unit.splitlines()), "source": reference, "players": []}
            for field, column in staff_columns.items():
                if column is None:
                    continue
                value = source_text(row[column])
                if field.endswith("_phone"):
                    value = "、".join(line.strip() for line in value.splitlines() if line.strip())
                team[field] = "" if value in ("无", "暂无") else value
            teams.append(team)
        elif any(column is not None and source_text(row[column]) for column in staff_columns.values()):
            raise ValueError(f"{reference} 的队务资料需要核对所属队伍")
        if not name:
            if any(value is not None and source_text(value) for value in row):
                raise ValueError(f"{reference} 缺少运动员姓名")
            continue
        if team is None:
            raise ValueError(f"{reference} 缺少所属单位")
        player = {"name": name, "source": reference}
        if number_column is not None:
            player["number"] = source_text(row[number_column])
        team["players"].append(player)
    return teams


def load_manifest(path, preset, protected, outputs):
    specs = array(load_json(path).get("sources"), "sources")
    inputs = []
    for spec in specs:
        formats = set(spec) & {"xlsx", "data"}
        if len(formats) != 1:
            raise ValueError("每个来源须指定一个 xlsx 或 data")
        kind = formats.pop()
        source = path.parent / required_text(spec[kind], "来源文件路径")
        sheets = spec.get("sheets", [])
        if not isinstance(sheets, list) or any(not isinstance(sheet, str) for sheet in sheets):
            raise ValueError(f"{source.name} 的 sheets 须为 工作表=组别 字符串数组")
        if kind == "data" and sheets:
            raise ValueError(f"{source.name} 的 sheets 只用于 xlsx")
        inputs.append((kind, source, sheets))
    separate_paths([path, *protected, *(source for _, source, _ in inputs)], outputs)

    groups = {}
    teams = {}
    for kind, source, sheets in inputs:
        extracted = (load_excel_groups(source, sheets, preset) if kind == "xlsx"
                     else array(load_json(source).get("groups"), f"{source.name} / groups"))
        for group in extracted:
            title = required_text(group.get("title"), f"{source.name} / 组别")
            merged = groups.setdefault(title, {"title": title, "teams": []})
            for team in array(group.get("teams"), f"{source.name} / {title}"):
                name = required_text(team.get("name"), f"{source.name} / 队伍")
                key = (title, name)
                if key not in teams:
                    teams[key] = {"name": name, "players": [], "field_sources": {}}
                    merged["teams"].append(teams[key])
                target = teams[key]
                reference = team.get("source") or source.name
                details = (*DETAILS, "liaison", "liaison_phone", "medic_phone") if preset == "inline-roster" else DETAILS
                for field in details:
                    value = source_text(team.get(field))
                    if not value:
                        continue
                    previous = target.get(field)
                    if previous and re.sub(r"\s+", "", previous) != re.sub(r"\s+", "", value):
                        previous_source = target["field_sources"][field]
                        raise ValueError(f"{title} / {name} 的 {field} 有冲突，请核对 {previous_source} 和 {reference}")
                    if not previous:
                        target[field] = value
                        target["field_sources"][field] = reference
                for person in array(team.get("players"), f"{reference} / {name} / players"):
                    target["players"].append({**person, "source": person.get("source") or reference})
    return list(groups.values())


def fill(element, fields):
    for node in element.iter(W + "t"):
        marker = MARKER.fullmatch(node.text or "")
        if marker:
            key = marker.group(1)
            if key not in fields:
                raise ValueError(f"母版字段未配置：{key}")
            node.text = fields[key]
    return element


def format_name(value):
    name = re.sub(r"\s+", "", value)
    return "  ".join(name) if len(name) == 2 else name


def make_team(team, team_index):
    label = required_text(team.get("name"), "队伍名称")
    fields = {"team": label}
    people = []
    for key in DETAILS:
        value = source_text(team.get(key))
        value = required_text(value, f"{label} / {key}") if value else ""
        if key in STAFF and value:
            reference = team.get("field_sources", {}).get(key, team.get("source"))
            people.append({"name": re.sub(r"\s+", "", value), "slot": key, "source": reference})
            value = format_name(value)
        fields[key] = value
    if fields["affiliation"] and not fields["affiliation"].startswith(("（", "(")):
        fields["affiliation"] = "（" + fields["affiliation"] + "）"

    players = array(team.get("players"), f"{label} / players")
    if not players:
        raise ValueError(f"{label} 没有运动员")
    names = [required_text(player.get("name"), f"{label} / 第{i}位运动员姓名")
             for i, player in enumerate(players, 1)]
    row_count = (len(players) + 1) // 2
    numbers = []
    for i, (player, name) in enumerate(zip(players, names), 1):
        formatted = format_name(name)
        number = source_text(player.get("number"))
        if number:
            number = required_text(number, f"{label} / 第{i}位运动员号码")
        row, column = (i, 2) if i <= row_count else (i - row_count, 5)
        numbers.append(number)
        fields[f"player_{i}"] = formatted
        people.append({
            "name": re.sub(r"\s+", "", name), "slot": "player", "order": i,
            "row": row, "column": column, "source": player.get("source"),
        })
    return fields, numbers, {"team": label, "order": team_index, "people": people}


def football_team(prototype, fields, numbers):
    team, lead, coach, assist, assistants, medic, table, gap = map(deepcopy, prototype)
    if any(fields[key] for key in ("coach_2", "affiliation")):
        raise ValueError(f"{fields['team']} 含双栏编号表未配置的教练或联队说明字段")
    for index in range(1, 4):
        fields[f"assistant_label_{index}"] = f"助理教练{index}：" if fields[f"assistant_{index}"] else ""
    staff = [lead, coach, assistants, medic] if fields["assistant_2"] or fields["assistant_3"] else [lead, coach, assist]
    row_count = (len(numbers) + 1) // 2
    header, row_template = table.findall(W + "tr")
    header_properties = header.find(W + "trPr")
    if header_properties.find(W + "tblHeader") is None:
        etree.SubElement(header_properties, W + "tblHeader")
    table.remove(row_template)
    for i in range(row_count):
        values = []
        for index in (i, i + row_count):
            if index < len(numbers):
                values.extend((str(index + 1), fields[f"player_{index+1}"], numbers[index]))
            else:
                values.extend(("", "", ""))
        row_fields = {f"cell_{index}": value for index, value in enumerate(values, 1)}
        row = fill(deepcopy(row_template), row_fields)
        if row_count > 8 or i == row_count - 1:
            for flag in row.iter(W + "keepNext"):
                flag.set(W + "val", "0")
        for height in row.iter(W + "trHeight"):
            height.set(W + "hRule", "atLeast")
        table.append(row)
    return [fill(p, fields) for p in (team, *staff, table, gap)]


def render_pages(args):
    if args.output.suffix.lower() != ".docx":
        raise ValueError("输出文件须使用 .docx 后缀")
    if args.sheet and not args.xlsx:
        raise ValueError("--sheet 只用于 --xlsx")
    preset = TEAM_ALIASES.get(args.preset, args.preset)
    if (preset == "inline-roster" or preset in INLINE_ALIASES) and (args.xlsx or args.manifest):
        protected = list(ASSETS.rglob("*.docx"))
        outputs = [args.output, args.audit]
        if args.xlsx:
            separate_paths([args.xlsx, *protected], outputs)
            groups = load_excel_groups(args.xlsx, args.sheet, "inline-roster")
        else:
            groups = load_manifest(args.manifest, "inline-roster", protected, outputs)
        render_sections(args, groups=groups)
        return
    if preset not in TEAM_PRESETS:
        if not args.data:
            raise ValueError("此母版请用 --data 提供分段名单 JSON，字段见内置母版目录")
        render_sections(args)
        return

    if args.data and len(args.data) != 1:
        raise ValueError("此母版的多份名单 JSON 请使用 --manifest")
    asset = ASSETS / f"{preset}.docx"
    source = args.xlsx or args.manifest or args.data[0]
    outputs = [args.output, args.audit]
    protected = list(ASSETS.rglob("*.docx"))
    separate_paths([*protected, source], outputs)
    if args.xlsx:
        groups = load_excel_groups(source, args.sheet, preset)
    elif args.manifest:
        groups = load_manifest(source, preset, protected, outputs)
    else:
        groups = array(load_json(source).get("groups"), "groups")
    if not groups:
        raise ValueError("没有组别")

    report = {"preset": args.preset, "groups": []}
    with ZipFile(asset) as template:
        document = parse_xml(template.read(BODY))
        body = document.find(W + "body")
        title, group_title, intro, *prototype, section = list(body)
        for element in list(body):
            body.remove(element)
        body.append(deepcopy(title))
        for index, group in enumerate(groups):
            group_name = required_text(group.get("title"), "组别")
            teams = array(group.get("teams"), group_name)
            if not teams:
                raise ValueError(f"{group_name} 没有队伍")
            heading = fill(deepcopy(group_title), {"group": group_name})
            if index:
                properties = heading.find(W + "pPr")
                page_break = properties.find(W + "pageBreakBefore")
                if page_break is None:
                    page_break = etree.Element(W + "pageBreakBefore")
                    keep_next = properties.find(W + "keepNext")
                    properties.insert(properties.index(keep_next) + 1 if keep_next is not None else 0, page_break)
                page_break.set(W + "val", "1")
            body.append(heading)
            if index == 0:
                body.append(deepcopy(intro))
            result = {"title": group_name, "teams": []}
            for order, team in enumerate(teams, 1):
                fields, numbers, entry = make_team(team, order)
                body.extend(football_team(prototype, fields, numbers))
                result["teams"].append(entry)
            body.remove(body[-1])
            report["groups"].append(result)
        body.append(section)
        xml = etree.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
        with temporary_output(args.output) as staged_docx, temporary_output(args.audit) as staged_audit:
            with ZipFile(staged_docx, "w") as output:
                for part in template.infolist():
                    output.writestr(part, xml if part.filename == BODY else template.read(part))
            save_json(staged_audit, report)
            staged_docx.replace(args.output)
            staged_audit.replace(args.audit)
    print(f"已生成 {sum(len(group['teams']) for group in report['groups'])} 支队伍的独立名单页")
