# /// script
# requires-python = ">=3.10"
# dependencies = ["docxtpl==0.20.2", "openpyxl==3.1.5", "python-docx>=1.2,<2", "lxml>=5,<7"]
# ///
"""Extract roster records and fill approved DOCX layouts."""

import argparse
import sys
import json
import re
from collections import Counter
from contextlib import contextmanager
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from uuid import uuid4
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined, TemplateError
from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


FIELDS = {"team", "role", "name"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
BODY = "word/document.xml"


def load_json(path):
    value = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(value, dict):
        raise ValueError(f"{path} 的 JSON 顶层必须为对象")
    return value


def array(value, label):
    if not isinstance(value, list) or any(not isinstance(item, dict) for item in value):
        raise ValueError(f"{label} 必须是对象数组")
    return value


def required_text(value, label):
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{label} 缺少文本")
    value = value.strip()
    if any(ord(char) < 32 or 0xD800 <= ord(char) <= 0xDFFF or ord(char) in (0xFFFE, 0xFFFF)
           for char in value):
        raise ValueError(f"{label} 包含换行、制表符或无效字符")
    return value


def integer(value, label, minimum=1):
    if type(value) is not int or value < minimum:
        raise ValueError(f"{label} 必须是大于等于 {minimum} 的整数")
    return value


@contextmanager
def temporary_output(path):
    with NamedTemporaryFile(dir=path.parent, delete=False) as handle:
        temporary_path = Path(handle.name)
    try:
        yield temporary_path
    finally:
        temporary_path.unlink(missing_ok=True)


def save_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def separate_paths(inputs, outputs):
    input_paths = {path.resolve() for path in inputs}
    output_paths = {path.resolve() for path in outputs}
    if input_paths & output_paths or len(output_paths) != len(outputs):
        raise ValueError("输入、输出和报告路径必须分开")


def excel_values(path, spec, header_row, last_row):
    book = load_workbook(path, data_only=True)
    try:
        sheet_name = required_text(spec.get("sheet"), f"{path.name} 的 sheet")
        if sheet_name not in book.sheetnames:
            raise ValueError(f"{path.name} 中没有工作表 {sheet_name}")
        sheet = book[sheet_name]
        end = last_row if last_row is not None else sheet.max_row
        if end > sheet.max_row:
            raise ValueError(f"{path.name} / {sheet_name} 的 last_data_row 超出工作表范围")
        values = [list(row) for row in sheet.iter_rows(
            min_row=header_row, max_row=end, values_only=True
        )]
        group_headers = {header for field, header in spec.get("columns", {}).items()
                         if field in {"team", "role"} and isinstance(header, str)}
        for region in sheet.merged_cells.ranges:
            # A single-column vertical merge commonly labels a block of people.
            if region.min_col != region.max_col or region.min_row <= header_row:
                continue
            header = sheet.cell(header_row, region.min_col).value
            if not isinstance(header, str) or header.strip() not in group_headers:
                continue
            label = sheet.cell(region.min_row, region.min_col).value
            for row_number in range(region.min_row + 1, min(region.max_row, end) + 1):
                values[row_number - header_row][region.min_col - 1] = label
        return sheet_name, values
    finally:
        book.close()


def word_values(path, spec, header_row, last_row):
    tables = Document(path).tables
    table_number = integer(spec.get("table"), f"{path.name} 的 table")
    if table_number > len(tables):
        raise ValueError(f"{path.name} 中没有第 {table_number} 张表")
    rows = tables[table_number - 1].rows
    end = last_row if last_row is not None else len(rows)
    if end > len(rows):
        raise ValueError(f"{path.name} / 表{table_number} 的 last_data_row 超出表格范围")
    values = [
        [None] * row.grid_cols_before + [cell.text for cell in row.cells]
        for row in rows[header_row - 1:end]
    ]
    return f"表{table_number}", values


def extract_source(config_path, spec):
    path = config_path.parent / required_text(spec.get("path"), "来源 path")
    header_row = integer(spec.get("header_row", 1), f"{path.name} 的 header_row")
    first = integer(spec.get("first_data_row", header_row + 1), "first_data_row", header_row + 1)
    last = spec.get("last_data_row")
    if last is not None:
        last = integer(last, "last_data_row", first)

    columns = spec.get("columns", {})
    constants = spec.get("constants", {})
    if not isinstance(columns, dict) or not isinstance(constants, dict):
        raise ValueError(f"{path.name} 的 columns 和 constants 必须为对象")
    if ("name" not in columns or set(columns) & set(constants)
            or set(columns) | set(constants) != FIELDS):
        raise ValueError(f"{path.name} 必须各映射一次 team、role、name，name 必须来自列")
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        location, rows = excel_values(path, spec, header_row, last)
    elif path.suffix.lower() == ".docx":
        location, rows = word_values(path, spec, header_row, last)
    else:
        raise ValueError(f"不支持直接提取 {path.name}，请转成统一名单 JSON")
    origin = f"{path.name} / {location}"
    if not rows:
        raise ValueError(f"{origin} 缺少表头行")
    headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    indices = {}
    for field, header in columns.items():
        header = required_text(header, f"{origin} 的 {field} 表头")
        if headers.count(header) != 1:
            raise ValueError(f"{origin} 的表头 {header!r} 必须出现一次")
        indices[field] = headers.index(header)
    if len(set(indices.values())) != len(indices):
        raise ValueError(f"{origin} 的多个字段指向同一列")

    for row_number, row in enumerate(rows[first - header_row:], first):
        if all(value is None or not str(value).strip() for value in row):
            continue
        source = f"{origin} / 第{row_number}行"
        values = {field: row[index] if index < len(row) else None
                  for field, index in indices.items()}
        values.update(constants)
        yield {
            **{field: required_text(values[field], f"{source} 的 {field}")
               for field in ("team", "role", "name")},
            "source": source,
        }


def extract(args):
    specs = array(load_json(args.config).get("sources"), "sources")
    separate_paths([args.config, *(args.config.parent / required_text(spec.get("path"), "path")
                                   for spec in specs)], [args.output])
    people = [person for spec in specs for person in extract_source(args.config, spec)]
    if not people:
        raise ValueError("没有提取到人员")
    with temporary_output(args.output) as staged:
        save_json(staged, {"people": people})
        staged.replace(args.output)
    print(f"已提取 {len(people)} 人")


def read_people(paths):
    people = []
    pending = []
    for path in paths:
        for number, raw in enumerate(array(load_json(path).get("people"), f"{path} 的 people"), 1):
            label = f"{path.name} / 第{number}条"
            person = {field: required_text(raw.get(field), f"{label} 的 {field}")
                      for field in ("team", "role", "name", "source")}
            review = raw.get("needs_review", False)
            if type(review) is not bool:
                raise ValueError(f"{label} 的 needs_review 必须是布尔值")
            if review:
                pending.append(f"{person['name']}（{person['source']}）")
            people.append(person)
    if pending:
        raise ValueError("请复核以下姓名：" + "、".join(pending))
    if not people:
        raise ValueError("名单中没有人员")
    return people


def section_plan(spec, people, marker_prefix):
    key = spec["id"]
    columns = integer(spec.get("columns"), f"{key} 的 columns")
    direction = spec.get("direction")
    if direction not in ("row", "column"):
        raise ValueError(f"{key} 的 direction 必须为 row 或 column")
    mode = spec.get("mode")
    if mode == "fixed":
        rows = integer(spec.get("rows"), f"{key} 的 rows")
        capacity = rows
    elif mode == "repeat":
        minimum = integer(spec.get("min_rows", 0), f"{key} 的 min_rows", 0)
        capacity = integer(spec.get("max_rows"), f"{key} 的 max_rows", max(1, minimum))
        rows = max(minimum, (len(people) + columns - 1) // columns)
    else:
        raise ValueError(f"{key} 的 mode 必须为 fixed 或 repeat")

    if len(people) > capacity * columns:
        raise ValueError(f"{key} 有 {len(people)} 人，母版容量为 {capacity * columns}，请按续页规则调整母版")
    name_limit = spec.get("max_name_chars")
    if name_limit is not None:
        name_limit = integer(name_limit, f"{key} 的 max_name_chars")
        for person in people:
            if len(person["name"]) > name_limit:
                raise ValueError(f"{key} 的姓名 {person['name']} 超过长度限制（{person['source']}）")

    matrix = [[f"{marker_prefix}{key}_{r + 1}_{c + 1}END" for c in range(columns)]
              for r in range(rows)]
    markers = dict.fromkeys((marker for row in matrix for marker in row), "")
    positions = [(r, c) for r in range(rows) for c in range(columns)]
    if direction == "column":
        positions.sort(key=lambda position: (position[1], position[0]))
    placements = []
    for person, (row, column) in zip(people, positions):
        markers[matrix[row][column]] = person["name"]
        placements.append({**person, "row": row + 1, "column": column + 1})
    if mode == "fixed":
        context = {f"r{r + 1}c{c + 1}": marker
                   for r, row in enumerate(matrix) for c, marker in enumerate(row)}
    else:
        context = [{f"c{c + 1}": marker for c, marker in enumerate(row)} for row in matrix]
    return context, markers, {"id": key, "rows": rows, "columns": columns,
                              "count": len(people), "placements": placements}


def plan_layout(layout_path, people):
    specs = array(load_json(layout_path).get("sections"), "sections")
    routes = {}
    ids = set()
    for spec in specs:
        key = required_text(spec.get("id"), "section id")
        if key != spec["id"]:
            raise ValueError(f"区块 id 两端不能有空白：{spec['id']!r}")
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", key) or key in ids:
            raise ValueError(f"区块 id 重复或无效：{key}")
        route = (required_text(spec.get("team"), f"{key} 的 team"),
                 required_text(spec.get("role"), f"{key} 的 role"))
        if route in routes:
            raise ValueError(f"队伍和角色重复配置：{route}")
        ids.add(key)
        routes[route] = (spec, [])
    for person in people:
        route = (person["team"], person["role"])
        if route not in routes:
            raise ValueError(f"未配置队伍与角色 {route}（{person['source']}）")
        routes[route][1].append(person)

    context = {}
    markers = {}
    sections = []
    prefix = f"ROSTER{uuid4().hex}SLOT"
    for spec, members in routes.values():
        value, slots, section = section_plan(spec, members, prefix)
        context[spec["id"]] = value
        markers.update(slots)
        sections.append(section)
    return context, markers, sections, prefix


def parse_xml(data):
    return etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True))


def xml_shape(element):
    return (element.tag, dict(element.attrib), element.text if element.text and element.text.strip() else "",
            [xml_shape(child) for child in element])


def layout_shapes(body):
    return [[xml_shape(node) for node in body.iter(W + tag)]
            for tag in ("sectPr", "tblGrid", "tblPr")]


def check_template_layout(before, after, original_body, rendered_body):
    if layout_shapes(original_body) != layout_shapes(rendered_body):
        raise ValueError("渲染改变了分节设置、表格属性或表格列宽")
    for part in before.namelist():
        if not re.fullmatch(r"word/(styles(?:WithEffects)?|settings|numbering|fontTable|header\d*|footer\d*)\.xml"
                            r"|word/theme/[^/]+\.xml", part):
            continue
        if part not in after.namelist() or xml_shape(parse_xml(before.read(part))) != xml_shape(
            parse_xml(after.read(part))
        ):
            raise ValueError(f"渲染改变了模板部件 {part}")


def fill_template(template_path, destination, context, markers, prefix):
    template = DocxTemplate(template_path)
    environment = Environment(undefined=StrictUndefined, autoescape=True)
    variables = template.get_undeclared_template_variables(jinja_env=environment)
    if variables != set(context):
        raise ValueError(f"模板区块与配置不符，模板缺少 {sorted(set(context) - variables)}，"
                         f"配置缺少 {sorted(variables - set(context))}")
    template.render(context, jinja_env=environment)
    stream = BytesIO()
    template.save(stream)
    with ZipFile(template_path) as before, ZipFile(stream) as after:
        body = parse_xml(after.read(BODY))
        check_template_layout(before, after, parse_xml(before.read(BODY)), body)
        seen = Counter()
        for node in body.iter(W + "t"):
            value = node.text or ""
            if prefix not in value:
                continue
            if value not in markers:
                raise ValueError("名单占位符必须单独位于一个 Word 文字片段中")
            seen[value] += 1
            node.text = markers[value]
            if node.text.startswith(" ") or node.text.endswith(" "):
                node.set(XML_SPACE, "preserve")
        if seen != Counter(markers.keys()):
            raise ValueError(f"格位数量不符，缺失 {len(markers.keys() - seen.keys())} 个，"
                             f"重复 {sum(count - 1 for count in seen.values())} 次")
        with ZipFile(destination, "w") as output:
            for entry in after.infolist():
                value = etree.tostring(body, encoding="UTF-8", xml_declaration=True, standalone=True) \
                    if entry.filename == BODY else after.read(entry)
                output.writestr(entry, value)


def render(args):
    if args.output.suffix.lower() != ".docx":
        raise ValueError("输出文件必须使用 .docx 后缀")
    separate_paths([args.template, args.layout, *args.roster], [args.output, args.audit])
    people = read_people(args.roster)
    context, markers, sections, prefix = plan_layout(args.layout, people)
    duplicates = Counter((p["team"], p["role"], p["name"]) for p in people)
    report = {"count": len(people), "sections": sections, "same_name": [
        {"team": team, "role": role, "name": name, "count": count,
         "sources": [p["source"] for p in people if (p["team"], p["role"], p["name"]) == (team, role, name)]}
        for (team, role, name), count in duplicates.items() if count > 1
    ]}
    with temporary_output(args.output) as staged_docx, temporary_output(args.audit) as staged_audit:
        fill_template(args.template, staged_docx, context, markers, prefix)
        save_json(staged_audit, report)
        staged_docx.replace(args.output)
        staged_audit.replace(args.audit)
    print(f"已填入 {len(people)} 人，覆盖 {len(sections)} 个区块")


def main():
    from roster_pages import add_command

    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    extraction = commands.add_parser("extract", help="从 Excel 或 Word 表格提取名单")
    extraction.add_argument("--config", type=Path, required=True)
    extraction.add_argument("--output", type=Path, required=True)
    extraction.set_defaults(run=extract)
    rendering = commands.add_parser("render", help="按 Word 母版排位并输出核对报告")
    rendering.add_argument("--template", type=Path, required=True)
    rendering.add_argument("--layout", type=Path, required=True)
    rendering.add_argument("--roster", type=Path, action="append", required=True)
    rendering.add_argument("--output", type=Path, required=True)
    rendering.add_argument("--audit", type=Path, required=True)
    rendering.set_defaults(run=render)
    add_command(commands)
    arguments = sys.argv[1:]
    # Keep invocations saved before the built-in catalog was unified working.
    if arguments[:1] == ["imported"]:
        arguments[0] = "pages"
    args = parser.parse_args(arguments)
    try:
        args.run(args)
    except (OSError, ValueError, TemplateError, BadZipFile, PackageNotFoundError,
            InvalidFileException, etree.XMLSyntaxError) as error:
        parser.exit(1, f"生成失败：{error}\n")


if __name__ == "__main__":
    main()
