"""Layout-aware structure shared by every reader.

The model is good at reading one clean table. It is bad at guessing where a
table starts, which row is the header, or that the row it just read continues on
the next page. This module does that bookkeeping deterministically:

* detect a header row from synonyms instead of an exact schema;
* carry title and multi-row header context;
* read a PDF whose roster is laid out one person per column (a transposed
  table), joining each person's fields across page breaks;
* give page batches their neighbouring text so a page is never read alone.

Nothing here calls a model. The callers decide which evidence to send.
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path

import pymupdf

# ---------------------------------------------------------------------------
# Header vocabulary
# ---------------------------------------------------------------------------

ALIASES = {
    "unit": ("参赛单位", "单位名称", "工作单位", "所在单位", "单位", "队名", "队伍名称",
             "参赛队伍", "代表队", "球队", "学校", "班级", "部门"),
    "team": ("参赛单位", "单位名称", "队名", "队伍名称", "参赛队伍", "代表队", "球队", "单位"),
    "group": ("参赛组别", "项目组别", "组别", "分组", "组名", "组"),
    "name": ("运动员姓名", "队员姓名", "参赛人员", "参赛运动员", "人员姓名", "学生姓名",
             "选手姓名", "联系人姓名", "姓名", "名字", "队员", "运动员"),
    "number": ("球衣号码", "参赛号码", "运动员号码", "队服号码", "号码牌", "球衣号",
               "球衣", "号码", "背号"),
    "gender": ("性别",),
    "id_card": ("身份证号码", "身份证号", "证件号码", "证件号", "身份证"),
    "phone": ("领队电话", "教练电话", "联系人电话", "联系电话", "联系方式", "手机号码",
              "手机号", "电话号码", "电话", "手机"),
    "leader": ("领队姓名", "领队", "联络员", "联系人", "负责人"),
    "coach": ("主教练", "教练员", "教练", "指导教师", "指导老师"),
    "event": ("比赛项目", "参赛项目", "竞赛项目", "项目"),
    "sign": ("签号", "序号", "编号", "道次", "签位"),
    "age": ("年龄",),
    "birth": ("出生日期", "出生年月", "出生时间", "生日", "出生"),
    "note": ("备注信息", "备注", "说明"),
    "contact": ("联系人", "联络人", "联系人姓名"),
    "role": ("岗位", "职务", "身份", "角色"),
}

# Fields that make a row a credible table header.
CORE_ROLES = {"name", "unit", "team", "number", "id_card", "phone", "leader",
              "coach", "contact", "group", "event", "sign", "gender", "age", "birth"}

NON_NAMES = {
    "姓名", "名字", "队员", "运动员", "领队", "教练", "队医", "队务", "联系人",
    "备注", "单位", "队伍", "球队", "学校", "班级", "部门", "组别", "项目",
    "号码", "球衣", "序号", "签号", "性别", "年龄", "日期", "身份证", "电话",
    "以上", "以下", "说明", "合计", "小计", "总计",
}

LABEL_KEYWORDS = (
    "姓名", "电话", "联系", "单位", "队伍", "球队", "组别", "项目", "备注",
    "盖章", "报名", "参赛", "颜色", "上衣", "短裤", "球袜", "替补", "教练",
    "领队", "队医", "队务", "以上", "序号", "号码", "性别", "身份证", "出生",
    "日期", "年龄", "签号", "职务", "岗位", "名称", "等级", "资料",
)

PUNCTUATION = re.compile(r"[\s\u3000:：()（）\[\]【】、，,。；;·]+")
HAN_NAME = re.compile(r"^[\u4e00-\u9fff·]{2,6}$")
JERSEY = re.compile(r"^(\d{1,3})号$")
DATE = re.compile(r"^(?:\d{4}[年./-]\d{1,2}[月./-]\d{1,2}[日号]?|\d{1,2}[月./-]\d{1,2}[日号]|\d{8})$")
ID_CARD = re.compile(r"^\d{17}[\dXx]$")
PHONE = re.compile(r"^1\d{10}$")
YES_NO = re.compile(r"^[是否]$")
TEAM_LABEL = re.compile(
    r"(?:参赛队伍名称|参赛单位名称|参赛单位|单位名称|队伍名称|代表队名称|球队名称|参赛队名)"
    r"\s*[:：]\s*([^\n\r（(]{2,40})"
)


def norm(value) -> str:
    """Remove whitespace and padding that carry no meaning in a cell."""
    return re.sub(r"\s+", "", str(value or "")).strip()


def header_key(value) -> str:
    """Normalize a header cell for matching."""
    return PUNCTUATION.sub("", str(value or "")).lower()


def role_for(value) -> str | None:
    """Best field role for one header cell, or ``None``.

    Exact aliases win; substring matches handle ``运动员姓名（必填）``; a loose
    fuzzy match is kept for typos such as ``球衣号玛``.
    """
    cell = header_key(value)
    if not cell or len(cell) > 24:
        return None
    best = None
    for role, aliases in ALIASES.items():
        for alias in aliases:
            key = header_key(alias)
            if cell == key:
                return role
            if len(key) >= 2 and key in cell:
                if best is None or len(key) > best[1]:
                    best = (role, len(key))
    if best:
        return best[0]
    if 2 <= len(cell) <= 6:
        for role, aliases in ALIASES.items():
            for alias in aliases:
                key = header_key(alias)
                if len(key) >= 2 and SequenceMatcher(None, cell, key).ratio() >= 0.82:
                    return role
    return None


def column_roles(values) -> dict[str, list[int]]:
    roles: dict[str, list[int]] = {}
    for index, value in enumerate(values):
        role = role_for(value)
        if role:
            roles.setdefault(role, []).append(index)
    return roles


@dataclass
class Header:
    """One header row plus the rows that qualify it."""

    index: int
    values: list[str]
    names: list[str]
    roles: dict[str, list[int]]
    data_start: int
    title_rows: list[list[str]] = field(default_factory=list)
    group_rows: list[list[str]] = field(default_factory=list)

    @property
    def width(self):
        return max([len(self.values), *(max(cols) + 1 for cols in self.roles.values())], default=0)

    def column(self, *roles):
        """First column index for any of ``roles``, or ``None``."""
        for role in roles:
            cols = self.roles.get(role)
            if cols:
                return cols[0]
        return None

    def as_dict(self):
        return {"index": self.index, "values": self.values, "names": self.names,
                "roles": self.roles, "data_start": self.data_start,
                "title_rows": self.title_rows}


def _row_values(row):
    if isinstance(row, (list, tuple)):
        return [norm(value) for value in row]
    return [norm(row)]


def _header_score(values, roles):
    distinct = {role for role in roles if role in CORE_ROLES}
    if len(distinct) < 2:
        return -1
    if not (distinct & {"name", "unit", "team", "contact", "leader"}):
        return -1
    if not (distinct & {"name", "unit", "team", "number", "id_card", "phone",
                        "contact", "leader", "group", "event", "sign"}):
        return -1
    filled = sum(bool(value) for value in values)
    numeric = sum(bool(re.fullmatch(r"\d{1,6}", value)) for value in values if value)
    score = 4 * len(distinct) + (3 if "name" in distinct else 0) + (2 if "unit" in distinct or "team" in distinct else 0)
    score += 2 if distinct & {"number", "id_card", "phone"} else 0
    score += 1 if filled >= 3 else -2
    score -= 3 * numeric
    return score


def find_header(rows, *, limit: int = 15) -> Header | None:
    """Find the most likely header row in the first ``limit`` rows.

    Handles a title row above the header and a two-row header where the upper
    row only carries group labels. Returns ``None`` when no row looks tabular.
    """
    rows = [_row_values(row) for row in rows[:limit]]
    if not rows:
        return None
    best = None
    for index, values in enumerate(rows):
        roles = column_roles(values)
        score = _header_score(values, roles)
        if score > 0 and (best is None or score > best[0]):
            best = (score, index, roles)
    if best is None:
        return None
    _, index, roles = best
    values = rows[index]
    width = max([len(values), *(max(cols) + 1 for cols in roles.values())], default=0)
    names = []
    for column in range(width):
        name = values[column] if column < len(values) else ""
        if not name:
            # A merged group header above the field row often supplies the
            # missing name; keep the first non-empty value above it.
            for above in range(index - 1, max(-1, index - 4), -1):
                cell = rows[above][column] if column < len(rows[above]) else ""
                if cell:
                    name = cell
                    break
        names.append(name)
    title_rows, group_rows = [], []
    for above in range(index - 1, -1, -1):
        values_above = rows[above]
        filled = [value for value in values_above if value]
        if not filled:
            continue
        roles_above = column_roles(values_above)
        positions = [position for position, value in enumerate(values_above) if value]
        if len(filled) >= 2 and sum(role in CORE_ROLES for role in roles_above) >= 1:
            group_rows.insert(0, values_above)
        elif len(filled) == 1 and positions and positions[0] > 0 and len(filled[0]) <= 12:
            group_rows.insert(0, values_above)
        elif len(filled) <= 2 or max((len(value) for value in filled), default=0) >= 10:
            title_rows.insert(0, values_above)
        if len(title_rows) + len(group_rows) >= 4:
            break
    return Header(index=index, values=values, names=names, roles=roles,
                  data_start=index + 1, title_rows=title_rows, group_rows=group_rows)


def repeated_header(rows, header: Header, *, start: int = 0) -> set[int]:
    """Indices of header rows repeated inside a long table."""
    names = {norm(value) for value in header.names if value}
    roles = set(header.roles)
    found = set()
    for index in range(start, len(rows)):
        values = _row_values(rows[index])
        if {role for role in column_roles(values) if role in CORE_ROLES} & roles:
            hits = sum(norm(value) in names for value in values if value)
            if hits >= max(2, len(names) // 2):
                found.add(index)
    return found


def _preview_cell(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def sheet_preview(path, unit, *, limit: int = 25) -> list[list[str]]:
    """First rows of one worksheet with every merged range filled."""
    path = Path(path)
    index = unit["index"]
    if path.suffix.lower() in {".xls", ".et"}:
        import xlrd

        with xlrd.open_workbook(str(path), on_demand=True, formatting_info=True, ragged_rows=True) as book:
            sheet = book.sheet_by_index(index - 1)
            grid = [[_preview_cell(sheet.cell_value(row, column))
                     for column in range(min(sheet.ncols, 30))]
                    for row in range(min(sheet.nrows, limit))]
            for rlo, rhi, clo, chi in getattr(sheet, "merged_cells", []):
                if rlo >= limit:
                    break
                value = _preview_cell(sheet.cell_value(rlo, clo))
                for row in range(rlo, min(rhi, limit)):
                    for column in range(clo, min(chi, 30)):
                        if row < len(grid) and column < len(grid[row]):
                            grid[row][column] = value
            return grid
    from openpyxl import load_workbook

    book = load_workbook(path, data_only=True, keep_links=False)
    try:
        sheet = book.worksheets[index - 1]
        merged = {}
        for area in sheet.merged_cells.ranges:
            value = _preview_cell(sheet.cell(area.min_row, area.min_col).value)
            for row in range(area.min_row, min(area.max_row, limit) + 1):
                for column in range(area.min_col, area.max_col + 1):
                    merged[(row, column)] = value
        width = min(sheet.max_column or 1, 30)
        grid = []
        for row in range(1, min(sheet.max_row or 1, limit) + 1):
            values = []
            for column in range(1, width + 1):
                values.append(merged.get((row, column),
                                         _preview_cell(sheet.cell(row, column).value)))
            grid.append(values)
        return grid
    finally:
        book.close()


def sheet_header(path, unit) -> Header | None:
    """Header for one worksheet, from a merged-cell-aware preview."""
    try:
        return find_header(sheet_preview(path, unit))
    except Exception:
        return None


# ---------------------------------------------------------------------------
# PDF layout: one person per column, fields continuing across pages
# ---------------------------------------------------------------------------


@dataclass
class PdfPage:
    index: int
    text: str
    context: list[str]
    has_text: bool


@dataclass
class PdfRecord:
    name: str
    number: str = ""
    fields: dict[str, str] = field(default_factory=dict)
    start_page: int = 1
    end_page: int = 1
    column: int = 0
    ordinal: int = 0
    y: float = 0.0
    x: float = 0.0

    def anchor(self, prefix: str) -> str:
        pages = (f"第 {self.start_page} 页" if self.start_page == self.end_page
                 else f"第 {self.start_page}–{self.end_page} 页")
        return f"{prefix} / {pages} / 名单第 {self.ordinal} 条"

    def value(self, team: str = "") -> dict[str, str]:
        data = {"姓名": self.name}
        if self.number:
            data["号码"] = self.number
        if team:
            data["单位"] = team
        data.update({key: value for key, value in self.fields.items() if value})
        return data


@dataclass
class PdfRoster:
    pages: list[PdfPage]
    records: list[PdfRecord]
    team: str = ""
    transposed: bool = False

    def by_page(self, page: int) -> list[PdfRecord]:
        return [record for record in self.records if record.start_page == page]

    def page(self, page: int) -> PdfPage | None:
        return next((item for item in self.pages if item.index == page), None)


@dataclass
class _Cell:
    page: int
    row: int
    x0: float
    y0: float
    x1: float
    y1: float
    text: str


def is_label(value: str) -> bool:
    text = norm(value)
    if not text:
        return True
    if text in NON_NAMES:
        return True
    if len(text) > 18:
        return True
    return any(keyword in text for keyword in LABEL_KEYWORDS)


def is_name(value: str) -> bool:
    text = norm(value)
    if not HAN_NAME.fullmatch(text) or text in NON_NAMES:
        return False
    return not any(keyword in text for keyword in LABEL_KEYWORDS)


def is_jersey(value: str) -> bool:
    return bool(JERSEY.fullmatch(norm(value)))


def is_date(value: str) -> bool:
    return bool(DATE.fullmatch(norm(value)))


def is_id_card(value: str) -> bool:
    return bool(ID_CARD.fullmatch(norm(value)))


def is_phone(value: str) -> bool:
    return bool(PHONE.fullmatch(norm(value)))


def is_yes_no(value: str) -> bool:
    return bool(YES_NO.fullmatch(norm(value)))


def _cell_text(page, rect) -> str:
    try:
        return norm(page.get_textbox(pymupdf.Rect(rect)))
    except Exception:
        return ""


def _collect_page_cells(page, page_index: int) -> list[_Cell]:
    cells: list[_Cell] = []
    if not page.get_text("text").strip():
        return cells
    try:
        tables = page.find_tables().tables
    except Exception:
        tables = []
    for table in tables:
        try:
            rows = list(table.rows)
        except Exception:
            continue
        for row_index, row in enumerate(rows):
            for rect in row.cells:
                if rect is None:
                    continue
                text = _cell_text(page, rect)
                if not text:
                    continue
                cells.append(_Cell(page_index, row_index, *rect, text))
    return cells


def _word_cells(page, page_index: int) -> tuple[list[_Cell], list[tuple[float, float]]]:
    """Rebuild a table grid from word positions when the PDF has no table lines.

    The hard part is not clustering words; it is ignoring form labels, staff
    names and logos that sit between roster columns. Rows with the modal number
    of roster-like cells (usually 4: one per athlete) define the column centres.
    """
    words = [word for word in page.get_text("words", sort=True) if word[4].strip()]
    if len(words) < 2:
        return [], []

    def clusters(values, gap):
        groups = []
        for value in sorted(values):
            if not groups or value - groups[-1] > gap:
                groups.append(value)
            elif value > groups[-1]:
                groups[-1] = value
        return groups

    def roster_like(value):
        text = norm(value)
        return (is_name(text) or is_date(text) or is_id_card(text) or is_yes_no(text)
                or bool(re.fullmatch(r"\d{1,3}", text)))

    def rows():
        centres = clusters([(word[1] + word[3]) / 2 for word in words], 8)
        result = []
        for index, centre in enumerate(centres):
            top = centre - 12 if index == 0 else (centres[index - 1] + centre) / 2
            bottom = centre + 12 if index == len(centres) - 1 else (centre + centres[index + 1]) / 2
            row = sorted((word for word in words if top <= (word[1] + word[3]) / 2 <= bottom),
                         key=lambda word: word[0])
            if row:
                result.append(row)
        return result

    def row_cells(row):
        cells, current, right = [], [], None
        for word in row:
            if current and word[0] - right > 12:
                cells.append(current)
                current = []
            current.append(word)
            right = max(right or 0, word[2])
        if current:
            cells.append(current)
        return cells

    candidates = []
    for row in rows():
        centres = []
        for cell in row_cells(row):
            text = norm(" ".join(word[4] for word in cell))
            if roster_like(text):
                centres.append((cell[0][0] + cell[-1][2]) / 2)
        if len(centres) >= 2:
            candidates.append(centres)
    counts = {len(values): sum(len(item) == len(values) for item in candidates)
              for values in candidates}
    target = max(counts, key=counts.get) if counts else 0
    centres = [centre for values in candidates if len(values) == target for centre in values]
    if len(clusters(centres, 80)) < 2:
        centres = [(word[0] + word[2]) / 2 for word in words if roster_like(word[4])]
    x_centres = clusters(centres, 80)
    if len(x_centres) < 2:
        x_centres = clusters([(word[0] + word[2]) / 2 for word in words], 80)
    x_bounds = []
    for index, centre in enumerate(x_centres):
        left = centre - 60 if index == 0 else (x_centres[index - 1] + centre) / 2
        right = centre + 60 if index == len(x_centres) - 1 else (centre + x_centres[index + 1]) / 2
        x_bounds.append((left, right))
    if x_bounds:
        x_bounds[0] = (min(x_bounds[0][0], min(word[0] for word in words)), x_bounds[0][1])
        x_bounds[-1] = (x_bounds[-1][0], max(x_bounds[-1][1], max(word[2] for word in words)))

    cells: list[_Cell] = []
    for row_index, row in enumerate(rows()):
        for column, (left, right) in enumerate(x_bounds):
            chunk = [word for word in row if left <= (word[0] + word[2]) / 2 <= right]
            if not chunk:
                continue
            chunk.sort(key=lambda word: word[0])
            text = norm(" ".join(word[4] for word in chunk))
            if not text:
                continue
            cells.append(_Cell(page_index, row_index,
                               min(word[0] for word in chunk), min(word[1] for word in chunk),
                               max(word[2] for word in chunk), max(word[3] for word in chunk), text))
    counts = Counter(cell.row for cell in cells)
    cells = [cell for cell in cells
             if not (counts[cell.row] == 1 and re.fullmatch(r"\d{1,3}", cell.text))]
    return cells, x_bounds


def _dedupe_cells(cells: list[_Cell]) -> list[_Cell]:
    seen = set()
    result = []
    for cell in cells:
        key = (cell.page, round(cell.x0, 1), round(cell.y0, 1), cell.text)
        if key in seen:
            continue
        seen.add(key)
        result.append(cell)
    return result


def _bounds(table) -> list[tuple[float, float]]:
    edges = set()
    for row in table.rows:
        for rect in row.cells:
            if rect is None:
                continue
            edges.add(round(rect[0], 1))
            edges.add(round(rect[2], 1))
    ordered = sorted(edges)
    merged = []
    for edge in ordered:
        if not merged or edge - merged[-1] > 12:
            merged.append(edge)
    return list(zip(merged, merged[1:]))


def _bin_for(x: float, bounds) -> int | None:
    for index, (left, right) in enumerate(bounds):
        if left - 1 <= x <= right + 1:
            return index
    return None


def _streams(cells: list[_Cell], bounds) -> dict[int, list[_Cell]]:
    streams: dict[int, list[_Cell]] = {index: [] for index in range(len(bounds))}
    for cell in cells:
        bin_index = _bin_for((cell.x0 + cell.x1) / 2, bounds)
        if bin_index is not None:
            streams[bin_index].append(cell)
    for values in streams.values():
        values.sort(key=lambda cell: (cell.page, cell.y0, cell.x0))
        seen = set()
        unique = []
        for cell in values:
            key = (cell.page, round(cell.y0 / 4), cell.text)
            if key in seen:
                continue
            seen.add(key)
            unique.append(cell)
        values[:] = unique
    return streams


def _record_values(record: PdfRecord) -> bool:
    return bool(record.name and record.number)


def _group_stream(stream: list[_Cell], column: int, *, plain_numbers: bool = False,
                  plain_rows=None) -> list[PdfRecord]:
    records: list[PdfRecord] = []
    current: PdfRecord | None = None
    pending: dict[str, str] = {}

    def number_text(cell: _Cell) -> str | None:
        text = norm(cell.text)
        match = JERSEY.fullmatch(text)
        if match:
            return match.group(1)
        if (plain_numbers and re.fullmatch(r"\d{1,3}", text) and 1 <= int(text) <= 999
                and plain_rows is not None and plain_rows.get((cell.page, cell.row), 0) >= 2):
            return text
        return None

    def start(cell: _Cell):
        nonlocal current, pending
        current = PdfRecord(name=cell.text, start_page=cell.page, end_page=cell.page,
                            column=column, y=cell.y0, x=cell.x0)
        for key, value in pending.items():
            if key == "number":
                current.number = value
            else:
                current.fields.setdefault(key, value)
        pending = {}
        records.append(current)

    for index, cell in enumerate(stream):
        text = cell.text
        number = number_text(cell)
        if is_name(text):
            following = [number_text(item) for item in stream[index + 1:index + 5]]
            previous = [number_text(item) for item in stream[max(0, index - 3):index]]
            if any(following) or any(previous):
                start(cell)
                continue
        if current is None:
            if number:
                pending.setdefault("number", number)
            elif is_date(text):
                pending.setdefault("出生日期", text)
            elif is_id_card(text):
                pending.setdefault("身份证号", text)
            elif is_yes_no(text):
                pending.setdefault("清真", text)
            continue
        if number and not current.number:
            current.number = number
        elif is_date(text):
            current.fields.setdefault("出生日期", text)
        elif is_id_card(text):
            current.fields.setdefault("身份证号", text)
        elif is_yes_no(text):
            current.fields.setdefault("清真", text)
        elif is_phone(text):
            current.fields.setdefault("电话", text)
        if cell.page > current.end_page:
            current.end_page = cell.page
    return records


def _transposed(cells: list[_Cell], candidates) -> list[PdfRecord]:
    best: list[PdfRecord] = []
    plain_rows = Counter((cell.page, cell.row) for cell in cells
                         if re.fullmatch(r"\d{1,3}", cell.text))
    for plain_numbers in (False, True):
        found: list[PdfRecord] = []
        for bounds in candidates:
            if len(bounds) < 2:
                continue
            streams = _streams(cells, bounds)
            records = []
            for column, stream in streams.items():
                records.extend(_group_stream(stream, column, plain_numbers=plain_numbers,
                                             plain_rows=plain_rows))
            records = [record for record in records if _record_values(record)]
            if len(records) > len(found):
                found = records
        if len(found) > len(best):
            best = found
    return best


def _regular(cells: list[_Cell], tables, pages: list[PdfPage]) -> list[PdfRecord]:
    """Best-effort row-wise table reader for PDFs with a real header."""
    records: list[PdfRecord] = []
    header = None
    for page_index, table in tables:
        try:
            rows = [[norm(value or "") for value in row] for row in table.extract()]
            width = table.col_count
        except Exception:
            continue
        if not rows:
            continue
        found = find_header(rows, limit=4)
        if found:
            header = found
        elif header is not None and header.width != width:
            header = None
        if header is None:
            continue
        name_col = header.column("name")
        number_col = header.column("number")
        unit_col = header.column("unit", "team")
        if name_col is None:
            continue
        for row in rows[header.data_start:]:
            if len(row) <= name_col or not is_name(row[name_col]):
                continue
            record = PdfRecord(name=row[name_col], column=0,
                               start_page=page_index, end_page=page_index)
            if number_col is not None and number_col < len(row):
                number = JERSEY.fullmatch(norm(row[number_col]))
                record.number = number.group(1) if number else norm(row[number_col])
            if unit_col is not None and unit_col < len(row):
                record.fields["单位"] = norm(row[unit_col])
            records.append(record)
    return records


def _page_context(page, table_rects) -> list[str]:
    """Text outside table boxes; a title or form caption is worth keeping."""
    lines = []
    try:
        blocks = page.get_text("blocks", sort=True)
    except Exception:
        return lines
    for block in blocks:
        text = norm(block[4])
        if not text:
            continue
        rect = pymupdf.Rect(block[:4])
        if any(rect.intersects(table_rect) for table_rect in table_rects):
            continue
        lines.append(text[:200])
    return lines[:12]


def pdf_roster(path: Path) -> PdfRoster:
    """Read every text table and return cross-page person records.

    A scanned PDF has no text cells; the returned roster is empty and callers
    fall back to page images. A text PDF whose roster is one person per column
    is joined by column across page boundaries, so a bib number on the next
    page still belongs to the name above it.
    """
    path = Path(path)
    pages: list[PdfPage] = []
    tables = []
    all_cells: list[_Cell] = []
    candidates: list[list[tuple[float, float]]] = []
    team = ""
    with pymupdf.open(path) as document:
        for page_index, page in enumerate(document, 1):
            text = page.get_text("text", sort=True)
            has_text = bool(text.strip())
            table_rects = []
            page_tables = []
            if has_text:
                try:
                    page_tables = page.find_tables().tables
                except Exception:
                    page_tables = []
                for table in page_tables:
                    try:
                        table_rects.append(pymupdf.Rect(table.bbox))
                    except Exception:
                        continue
                    bounds = _bounds(table)
                    if 2 <= len(bounds) <= 10:
                        candidates.append(bounds)
                all_cells.extend(_collect_page_cells(page, page_index))
                word_cells, word_bounds = _word_cells(page, page_index)
                all_cells.extend(word_cells)
                if 2 <= len(word_bounds) <= 10:
                    candidates.append(word_bounds)
            context = _page_context(page, table_rects) if has_text else []
            if not context and page_index == 1 and has_text:
                context = [norm(line) for line in text.splitlines() if norm(line)][:40]
            if page_index == 1 and has_text:
                match = TEAM_LABEL.search(text)
                if match:
                    team = norm(match.group(1)).rstrip("，,。;；")
            pages.append(PdfPage(page_index, text, context, has_text))
            tables.extend((page_index, table) for table in page_tables)
    all_cells = _dedupe_cells(all_cells)
    transposed = _transposed(all_cells, candidates)
    regular = _regular(all_cells, tables, pages) if not transposed else []
    records = transposed or regular
    records.sort(key=lambda record: (record.start_page, round(record.y / 40), record.x))
    for ordinal, record in enumerate(records, 1):
        record.ordinal = ordinal
    return PdfRoster(pages=pages, records=records, team=team, transposed=bool(transposed))


def image_looks_tabular(image) -> bool:
    """Detect a scanned table from long continuous dark lines.

    Text-only pages have many short dark runs (glyphs); a form or roster page
    has long horizontal and vertical rules. This is a cheap raster check, not
    OCR, and is only used to stop an empty submission on a page that clearly
    contains a table.
    """
    from PIL import Image, ImageOps

    source = image if hasattr(image, "width") else Image.open(image)
    try:
        grey = ImageOps.grayscale(source)
    finally:
        if source is not image:
            source.close()
    width = 600
    height = max(1, int(grey.height * width / max(1, grey.width)))
    grey = grey.resize((width, height))
    pixels = list(grey.get_flattened_data() if hasattr(grey, "get_flattened_data") else grey.getdata())
    dark = [value < 200 for value in pixels]
    horizontal = vertical = 0
    for row in range(height):
        best = run = 0
        base = row * width
        for column in range(width):
            if dark[base + column]:
                run += 1
                best = max(best, run)
            else:
                run = 0
        if best > width * .5:
            horizontal += 1
    for column in range(width):
        best = run = 0
        for row in range(height):
            if dark[row * width + column]:
                run += 1
                best = max(best, run)
            else:
                run = 0
        if best > height * .5:
            vertical += 1
    return horizontal >= 4 and vertical >= 2


def page_looks_tabular(path, page_index: int) -> bool:
    from io import BytesIO

    with pymupdf.open(path) as document:
        page = document[page_index - 1]
        pixmap = page.get_pixmap(matrix=pymupdf.Matrix(1, 1), alpha=False)
        data = pixmap.tobytes("png")
    try:
        return image_looks_tabular(BytesIO(data))
    except Exception:
        return False


def pdf_page_notes(roster: PdfRoster, page: int) -> list[str]:
    """Short, page-local hints that tell the model how to read its neighbours."""
    item = roster.page(page)
    if item is None:
        return []
    notes = []
    if roster.records:
        if roster.by_page(page):
            if roster.transposed:
                notes.append("本页名单表是一人一列、字段跨页的版式；已按列合并，只提交姓名开始于本页的记录。")
            else:
                notes.append("本页名单已按人员整理；只提交姓名开始于本页的记录，相邻页仅用于补全字段。")
        else:
            notes.append("本页可能只是上一页名单的续页；若没有以本页姓名开头的记录，不要重复提交上一页人员。")
    if not item.has_text:
        notes.append("本页没有文字层，请查看页图；页图可能同时包含相邻页，只在当前页开始的人员计入本批。")
    return notes


def pdf_page_context(roster: PdfRoster, page: int) -> list[str]:
    item = roster.page(page)
    if item is None:
        return []
    context = list(item.context)
    if roster.team and page == 1:
        context.insert(0, f"参赛单位：{roster.team}")
    return context[:14]
