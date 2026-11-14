"""Size the columns of generated roster tables for their contents and page."""

from copy import deepcopy
import unicodedata

from lxml import etree

from roster import W


def child(parent, name):
    node = parent.find(W + name)
    return node if node is not None else etree.SubElement(parent, W + name)


def text_width(value, em=210):
    # Estimate 10.5 pt text in twips. Reserve a little room for cell padding.
    text = "" if value is None else str(value)
    return 160 + sum(em if unicodedata.east_asian_width(char) in "WF" else round(em * 0.55) for char in text)


def usable_width(section):
    page = section.find(W + "pgSz")
    margins = section.find(W + "pgMar")
    return int(page.get(W + "w")) - sum(int(margins.get(W + side)) for side in ("left", "right"))


def set_orientation(section, landscape):
    page = section.find(W + "pgSz")
    width, height = sorted(int(page.get(W + edge)) for edge in ("w", "h"))
    if landscape:
        width, height = height, width
        page.set(W + "orient", "landscape")
    else:
        page.attrib.pop(W + "orient", None)
    page.set(W + "w", str(width))
    page.set(W + "h", str(height))


def set_table_widths(table, widths):
    properties = table.find(W + "tblPr")
    child(properties, "tblLayout").set(W + "type", "fixed")
    indent = properties.find(W + "tblInd")
    if indent is not None:
        indent.set(W + "w", "0")
    child(properties, "tblW").attrib.update({W + "type": "dxa", W + "w": str(sum(widths))})
    grid = table.find(W + "tblGrid")
    grid[:] = []
    for width in widths:
        etree.SubElement(grid, W + "gridCol").set(W + "w", str(width))
    for row in table.iter(W + "tr"):
        cursor = 0
        for cell in row.findall(W + "tc"):
            properties = cell.find(W + "tcPr")
            span = properties.find(W + "gridSpan")
            count = 1 if span is None else int(span.get(W + "val"))
            width = sum(widths[cursor:cursor + count])
            child(properties, "tcW").attrib.update({W + "type": "dxa", W + "w": str(width)})
            cursor += count


def preferred_columns(table):
    columns = len(table.find(W + "tblGrid"))
    # Keep identifiers, names and headings legible when a long field uses
    # most of the remaining width. These are minimums, not preset widths.
    minimum = {
        3: [1100, 700, 1500],             # group, sequence, team
        4: [700, 1500, 1350, 1650],       # sequence, unit, contact, phone
        5: [700, 1250, 1080, 1450, 850],  # sequence, unit, name, event, sign
    }[columns]
    preferred = minimum.copy()
    for row in table.findall(W + "tr"):
        header = row.find(f"{W}trPr/{W}tblHeader") is not None
        cursor = 0
        for cell in row.findall(W + "tc"):
            span = cell.find(f"{W}tcPr/{W}gridSpan")
            if span is not None:
                cursor += int(span.get(W + "val"))
                continue
            value = "".join(node.text or "" for node in cell.iter(W + "t"))
            sizes = [int(n.get(W + "val")) for n in cell.iter(W + "sz")]
            em = max(sizes, default=21) * 10
            measured = min(text_width(value, em), 4300)
            if header:
                minimum[cursor] = max(minimum[cursor], measured + 100)
            preferred[cursor] = max(preferred[cursor], measured, minimum[cursor])
            cursor += 1
    return preferred, minimum


def distribute_widths(preferred, minimum, available):
    slack = available - sum(minimum)
    weights = [width - floor for width, floor in zip(preferred, minimum)]
    if not any(weights):
        weights = minimum
    total = sum(weights)
    widths = [floor + slack * weight // total for floor, weight in zip(minimum, weights)]
    widths[-1] += available - sum(widths)
    return widths


def adapt_regular_tables(content, section, contact_orientation=None):
    tables = [table for node in content for table in node.iter(W + "tbl")]
    measurements = [preferred_columns(table) for table in tables]
    if contact_orientation is not None:
        landscape = contact_orientation == "landscape" or (
            contact_orientation == "auto"
            and any(sum(preferred) > usable_width(section) for preferred, _ in measurements)
        )
        set_orientation(section, landscape)
    available = usable_width(section)
    for table, (preferred, minimum) in zip(tables, measurements):
        set_table_widths(table, distribute_widths(preferred, minimum, available))


def cap_horizontal_padding(cell, table_margins, text_width):
    properties = cell.find(W + "tcPr")
    width = int(properties.find(W + "tcW").get(W + "w"))
    limit = max(0, (width - text_width) // 2)
    margins = properties.find(W + "tcMar")
    for side in ("start", "end"):
        current = margins.find(W + side) if margins is not None else None
        inherited = current if current is not None else table_margins.find(W + side)
        if inherited is None or int(inherited.get(W + "w")) <= limit:
            continue
        if margins is None:
            margins = etree.Element(W + "tcMar")
            before = properties.find(W + "vAlign")
            position = len(properties) if before is None else properties.index(before)
            properties.insert(position, margins)
        child(margins, side).attrib.update({W + "type": "dxa", W + "w": str(limit)})


def prepare_team_table(table, section, teams, has_contacts, has_title):
    available = usable_width(section)
    metadata = 4 if has_contacts else 3
    team_width = max(1600, min(3400, max(text_width(t.get("team")) for t in teams)))
    leader_width = max(900, min(1400, max(text_width(t.get("leader")) for t in teams)))
    widths = [660, team_width, leader_width]
    if has_contacts:
        widths.append(max(1520, min(2400, max(text_width(t.get("phone")) for t in teams))))

    players = [t.get("players", []) for t in teams]
    longest = max((text_width(p.get("name")) for group in players for p in group), default=0)
    name_width = max(760, min(1250, longest))
    free = available - sum(widths)
    slots = max(1, min(max(map(len, players)), free // name_width))
    slot_width = min(1250, free // slots)
    widths.extend([slot_width] * slots)

    title, header = table.findall(W + "tr")
    if has_title:
        title.find(f"{W}tc/{W}tcPr/{W}gridSpan").set(W + "val", str(len(widths)))
    else:
        table.remove(title)
    if not has_contacts:
        header.remove(header.findall(W + "tc")[3])
    header.findall(W + "tc")[-1].find(f"{W}tcPr/{W}gridSpan").set(W + "val", str(slots))

    matrix = next(tag for tag in table.iter(W + "tag") if tag.get(W + "val").startswith("matrix:players:"))
    matrix.set(W + "val", f"matrix:players:{slots}")
    prototype = matrix.getparent().getparent().find(W + "sdtContent")
    first = prototype[0]
    source_cells = first.findall(W + "tc")
    player_cell = deepcopy(source_cells[4])
    if not has_contacts:
        first.remove(source_cells[3])
    for cell in source_cells[4:]:
        first.remove(cell)
    for cell in first.findall(W + "tc"):
        properties = cell.find(W + "tcPr")
        merge = properties.find(W + "vMerge")
        if merge is not None:
            properties.remove(merge)
    for index in range(1, slots + 1):
        cell = deepcopy(player_cell)
        next(cell.iter(W + "t")).text = f"{{{{player{index}?}}}}"
        first.append(cell)

    for row in list(prototype)[1:]:
        prototype.remove(row)
    if max(map(len, players)) > slots:
        continuation = deepcopy(first)
        for start, following in zip(first.findall(W + "tc")[:metadata],
                                    continuation.findall(W + "tc")[:metadata]):
            child(start.find(W + "tcPr"), "vMerge").set(W + "val", "restart")
            child(following.find(W + "tcPr"), "vMerge").attrib.clear()
            for text in following.iter(W + "t"):
                text.text = ""
        prototype.append(continuation)
    for row in prototype:
        for flag in row.iter(W + "keepNext"):
            flag.set(W + "val", "0")
    set_table_widths(table, widths)
    margins = table.find(f"{W}tblPr/{W}tblCellMar")
    for row in prototype:
        for cell in row.findall(W + "tc")[metadata:]:
            cap_horizontal_padding(cell, margins, 666)
    return widths
