import subprocess
from contextlib import contextmanager
from io import BytesIO
from itertools import islice
from pathlib import Path
from shutil import which
from struct import error as StructError
from tempfile import TemporaryDirectory
from zipfile import BadZipFile, ZipFile

import pymupdf
import xlrd

from limits import BATCH_ROWS
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from lxml import etree, html
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from PIL import Image, ImageOps

def find_soffice() -> str | None:
    """Locate the LibreOffice soffice binary across env overrides and macOS/Linux paths."""
    from os import getenv

    candidates = [
        getenv("LIBREOFFICE_BIN", "").strip(),
        getenv("SOFFICE_PATH", "").strip(),
        "soffice",
        "libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice",
        str(
            Path.home()
            / "Library/Application Support/TRAE SOLO CN/ModularData/ai-agent/vm/tools/opt/libreoffice/LibreOffice.app/Contents/MacOS/soffice"
        ),
    ]
    for candidate in candidates:
        if not candidate:
            continue
        found = which(candidate)
        if found:
            return found
    return None


LEGACY_SHEETS = {".xls", ".et"}
TEXTUTIL = Path("/usr/bin/textutil")
EXTENSIONS = {".xlsx", ".docx", ".pdf", ".png", ".jpg", ".jpeg", ".webp"} | LEGACY_SHEETS
SOFFICE = find_soffice()
if TEXTUTIL.is_file() or SOFFICE:
    EXTENSIONS.add(".wps")
if SOFFICE:
    EXTENSIONS.add(".doc")
Image.MAX_IMAGE_PIXELS = 24_000_000
BATCH = BATCH_ROWS
WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def value_text(cell):
    value = cell.value
    if value is None:
        return ""
    if type(value) is int and cell.number_format.strip("0") == "" and cell.number_format:
        return str(value).zfill(len(cell.number_format))
    return str(value)


@contextmanager
def legacy_workbook(path):
    try:
        with xlrd.open_workbook(str(path), on_demand=True, formatting_info=True, ragged_rows=True) as book:
            yield book
    except (xlrd.XLRDError, StructError, IndexError, AssertionError) as exc:
        if "encrypt" in str(exc).lower():
            raise ValueError("请先解密表格") from exc
        raise ValueError(f"无法读取 {path.suffix.upper()} 表格，请检查原件或在 Excel、WPS 中另存为 XLSX") from exc


def legacy_cell_text(book, cell):
    if cell.ctype in {xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK}:
        return ""
    value = cell.value
    if cell.ctype == xlrd.XL_CELL_NUMBER and value.is_integer():
        integer = int(value)
        number_format = book.format_map.get(book.xf_list[cell.xf_index].format_key)
        if number_format and number_format.format_str and not number_format.format_str.strip("0"):
            return str(integer).zfill(len(number_format.format_str))
        return str(integer)
    if cell.ctype == xlrd.XL_CELL_DATE:
        date = xlrd.xldate_as_datetime(value, book.datemode)
        if 0 <= value < 1:
            return date.time().isoformat()
        if value.is_integer():
            return date.date().isoformat()
        return date.isoformat(sep=" ")
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return str(bool(value))
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(value, "#ERROR")
    return str(value)


def legacy_sheet_rows(path, index, prefix, start, stop, *, include_merges=True):
    with legacy_workbook(path) as book:
        sheet = book.sheet_by_index(index - 1)
        merged = {}
        if include_merges:
            for rlo, rhi, clo, chi in getattr(sheet, "merged_cells", []):
                if clo + 1 != chi:
                    continue
                value = legacy_cell_text(book, sheet.cell(rlo, clo))
                for row in range(rlo + 1, rhi):
                    merged[(row, clo)] = value
        for i in range(start, min(stop, sheet.nrows) + 1):
            values = [legacy_cell_text(book, cell) for cell in sheet.row(i - 1)]
            for column, value in ((column, value) for (row, column), value in merged.items() if row == i - 1):
                if column < len(values):
                    values[column] = value
            yield f"{prefix} / {sheet.name} / 第 {i} 行", values


def wps_rows(path, prefix):
    if not TEXTUTIL.is_file():
        raise ValueError("当前服务器缺少 WPS 文字读取器，请将原件另存为 DOCX")
    try:
        result = subprocess.run(
            [str(TEXTUTIL), "-convert", "html", "-format", "doc", "-stdout", "-noload", "-nostore",
             "-strip", "--", str(path.resolve())],
            stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, timeout=30, check=True,
        )
        body = html.document_fromstring(result.stdout, parser=html.HTMLParser(no_network=True)).find("body")
    except subprocess.TimeoutExpired as exc:
        raise ValueError("读取 WPS 文字超时，请拆分文档或另存为 DOCX") from exc
    except (subprocess.CalledProcessError, etree.ParserError) as exc:
        raise ValueError("无法读取此 WPS 文字文档，请在 WPS 中另存为 DOCX 后重试") from exc
    if body is None:
        raise ValueError("WPS 文字文档缺少正文")

    def content(element):
        text = "".join(
            part if isinstance(part, str) else "\n"
            for part in element.xpath(".//text() | .//br")
        )
        return text.strip()

    def cell_content(cell):
        paragraphs = cell.findall("p")
        return "\n".join(content(part) for part in paragraphs) if paragraphs else content(cell)

    table = paragraph = 0
    for block in body:
        if block.tag == "table":
            table += 1
            for i, row in enumerate(block.xpath("./tr | ./*[self::thead or self::tbody or self::tfoot]/tr"), 1):
                cells = row.xpath("./td | ./th")
                yield f"{prefix} / 表 {table} / 第 {i} 行", [cell_content(cell) for cell in cells]
        elif block.tag in {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6"}:
            paragraph += 1
            yield f"{prefix} / 段落 {paragraph}", content(block)


def check_office(path, *, bulk=False):
    with ZipFile(path) as archive:
        entries = archive.infolist()
        if len(entries) > (10_000 if bulk else 2500) or sum(entry.file_size for entry in entries) > (512 if bulk else 80) * 1024**2:
            raise ValueError("Office 文件解压后超过 Demo 上限")
        if bulk and any(entry.filename.endswith((".xml", ".rels")) and entry.file_size > 32 * 1024**2
                        for entry in entries):
            raise ValueError("Office 文件含超过 32 MB 的单个结构文件，请拆分")
        if any(entry.flag_bits & 1 for entry in entries):
            raise ValueError("请先解密文件")


@contextmanager
def converted_doc(path):
    if not SOFFICE:
        raise ValueError("当前服务器缺少旧版 Word 读取器，请将原件另存为 DOCX")
    with TemporaryDirectory(prefix="roster-doc-") as directory:
        folder = Path(directory).resolve()
        output = folder / (path.stem + ".docx")
        try:
            subprocess.run(
                [SOFFICE, f"-env:UserInstallation={(folder / 'profile').as_uri()}",
                 "--headless", "--convert-to", "docx", "--outdir", str(folder), str(path.resolve())],
                stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                timeout=30, check=True,
            )
        except subprocess.TimeoutExpired as exc:
            raise ValueError("读取旧版 Word 超时，请另存为 DOCX 后重试") from exc
        except (OSError, subprocess.CalledProcessError) as exc:
            raise ValueError("无法转换旧版 Word，请另存为 DOCX 后重试") from exc
        if not output.is_file():
            raise ValueError("无法读取此 DOC 文件，请检查文件或另存为 DOCX")
        yield output


def inventory(path: Path, *, bulk=False):
    try:
        return _inventory(path, bulk=bulk)
    except (BadZipFile, PackageNotFoundError, InvalidFileException, etree.XMLSyntaxError,
            pymupdf.FileDataError, Image.DecompressionBombError) as exc:
        raise ValueError(f"无法读取 {path.suffix.upper()} 文件，请检查文件格式、完整性和大小。") from exc


def _inventory(path: Path, *, bulk=False):
    suffix = path.suffix
    if suffix == ".doc":
        with converted_doc(path) as converted:
            return inventory(converted, bulk=bulk)
    if suffix in {".xlsx", ".docx"}:
        check_office(path, bulk=bulk)
    if suffix in LEGACY_SHEETS:
        max_rows = 5000 if bulk else 700
        with legacy_workbook(path) as book:
            units = []
            for index in range(book.nsheets):
                sheet = book.sheet_by_index(index)
                if sheet.nrows > max_rows or sheet.ncols > 60:
                    raise ValueError(f"工作表「{sheet.name}」超过上限（{max_rows} 行 × 60 列）")
                units.append({"label": sheet.name, "size": sheet.nrows, "index": index + 1, "kind": "sheet"})
                if book.on_demand:
                    book.unload_sheet(index)
            return units
    if suffix == ".xlsx":
        book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            units = []
            for index, sheet in enumerate(book.worksheets, 1):
                rows = sheet.max_row
                max_rows = 5000 if bulk else 700
                if rows is None or rows > max_rows or (sheet.max_column or 0) > 60:
                    raise ValueError(f"工作表「{sheet.title}」超过上限（{max_rows} 行 × 60 列）或缺少行范围")
                units.append({"label": sheet.title, "size": rows, "index": index, "kind": "sheet"})
            return units
        finally:
            book.close()
    if suffix in {".docx", ".wps"}:
        if suffix == ".wps":
            count = sum(1 for _ in wps_rows(path, ""))
        elif bulk:
            count = sum(1 for _ in bulk_word_rows(path, ""))
        else:
            doc = Document(path)
            count = sum(len(block.rows) if hasattr(block, "rows") else 1 for block in doc.iter_inner_content())
        max_rows = 5000 if bulk else 800
        if count > max_rows:
            raise ValueError(f"Word 超过上限（{max_rows} 段或表格行）")
        label = "WPS 正文" if suffix == ".wps" else "Word 正文"
        return [{"label": label, "size": count, "index": 1, "kind": "word"}]
    if suffix == ".pdf":
        with pymupdf.open(path) as doc:
            if doc.needs_pass:
                raise ValueError("请先解密 PDF")
            max_pages = 300 if bulk else 25
            if len(doc) > max_pages:
                raise ValueError(f"PDF 超过上限（{max_pages} 页）")
            return [{"label": f"第 {i} 页", "size": 1, "index": i, "kind": "page"}
                    for i in range(1, len(doc) + 1)]
    with Image.open(path) as image:
        if image.width * image.height > Image.MAX_IMAGE_PIXELS:
            raise ValueError("图片像素过大")
        image.verify()
    return [{"label": "图片", "size": 1, "index": 1, "kind": "image"}]


def extract(path, unit, start, prefix, *, bulk=False, row_limit=None):
    if path.suffix == ".doc":
        with converted_doc(path) as converted:
            return extract(converted, unit, start, prefix, bulk=bulk, row_limit=row_limit)
    kind, index = unit["kind"], unit["index"]
    stop = min(unit["size"], start + (row_limit or BATCH) - 1)
    if kind == "sheet" and path.suffix in LEGACY_SHEETS:
        rows = list(legacy_sheet_rows(path, index, prefix, start, stop))
    elif kind == "sheet":
        book = load_workbook(path, data_only=True, keep_links=False)
        try:
            sheet = book.worksheets[index - 1]
            merged = {}
            for area in sheet.merged_cells.ranges:
                if area.min_col != area.max_col:
                    continue
                value = value_text(sheet.cell(area.min_row, area.min_col))
                for row_number in range(area.min_row + 1, area.max_row + 1):
                    merged.setdefault(row_number, {})[area.min_col - 1] = value
            rows = []
            for i, row in enumerate(sheet.iter_rows(min_row=start, max_row=stop), start):
                values = [value_text(c) for c in row]
                for column, value in merged.get(i, {}).items():
                    if column < len(values):
                        values[column] = value
                rows.append((f"{prefix} / {sheet.title} / 第 {i} 行", values))
        finally:
            book.close()
    elif kind == "word" and path.suffix == ".wps":
        rows = list(islice(wps_rows(path, prefix), start - 1, stop))
    elif kind == "word" and bulk:
        rows = list(islice(bulk_word_rows(path, prefix), start - 1, stop))
    elif kind == "word":
        rows = []
        tables = paragraphs = 0
        for block in Document(path).iter_inner_content():
            if hasattr(block, "rows"):
                tables += 1
                rows.extend((f"{prefix} / 表 {tables} / 第 {i} 行", [cell.text for cell in row.cells])
                            for i, row in enumerate(block.rows, 1))
            else:
                paragraphs += 1
                rows.append((f"{prefix} / 段落 {paragraphs}", block.text))
        rows = rows[start - 1:stop]
    elif kind == "page":
        with pymupdf.open(path) as doc:
            page = doc[index - 1]
            lines = page.get_text("text", sort=True).splitlines()
            mixed_page = bulk and page_has_large_image(page)
        stop = min(len(lines), start + (row_limit or BATCH) - 1)
        rows = [(f"{prefix} / 第 {index} 页 / 第 {i} 行", lines[i - 1])
                for i in range(start, stop + 1)]
        return rows, stop + 1 if stop < len(lines) else None, mixed_page or not any(lines), len(lines)
    else:
        return [], None, True, unit["size"]
    return rows, stop + 1 if stop < unit["size"] else None, False, unit["size"]


def page_has_large_image(page):
    area = page.rect.get_area()
    return any((pymupdf.Rect(item["bbox"]) & page.rect).get_area() >= area * .12
               for item in page.get_image_info())


def page_image(path, unit, prefix):
    if unit["kind"] == "page":
        pages = unit.get("image_pages") or [unit["index"]]
        pages = list(dict.fromkeys(int(page) for page in pages))[:3]
        with pymupdf.open(path) as doc:
            max_width = 2300 / max(1, len(pages))
            images = []
            for page_index in pages:
                page = doc[page_index - 1]
                scale = min(2.5, max_width / max(1, page.rect.width),
                            1800 / max(1, page.rect.height))
                data = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False).tobytes("png")
                images.append(Image.open(BytesIO(data)).convert("RGB"))
        if len(images) == 1:
            image = images[0]
        else:
            gap = 12
            width = sum(item.width for item in images) + gap * (len(images) - 1)
            height = max(item.height for item in images)
            image = Image.new("RGB", (width, height), "white")
            left = 0
            for label, item in zip(pages, images):
                image.paste(item, (left, 0))
                try:
                    from PIL import ImageDraw, ImageFont
                    font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 26)
                except Exception:
                    from PIL import ImageDraw, ImageFont
                    font = ImageFont.load_default()
                draw = ImageDraw.Draw(image)
                draw.rectangle((left + 4, 4, left + 44, 38), fill="white")
                draw.text((left + 10, 6), str(label), fill="black", font=font)
                left += item.width + gap
        anchor = (f"{prefix} / 第 {pages[0]} 页图" if len(pages) == 1
                  else f"{prefix} / 第 {pages[0]}–{pages[-1]} 页图")
    elif unit["kind"] == "image":
        anchor = f"{prefix} / 图片"
        image = Image.open(path)
    else:
        raise ValueError("只有 PDF 和图片可读取页图")
    with image:
        rgb = ImageOps.exif_transpose(image).convert("RGB")
        rgb = ImageOps.autocontrast(rgb, cutoff=1)
        if unit["kind"] == "image" and rgb.height > rgb.width * 1.8 and rgb.height > 1500:
            # A long phone screenshot read as one image loses half its text
            # resolution. Split it into two overlapping halves and place them
            # side by side, labelled so the model reads left then right.
            overlap = max(40, int(rgb.height * .06))
            middle = rgb.height // 2
            halves = [rgb.crop((0, 0, rgb.width, middle + overlap)),
                      rgb.crop((0, middle - overlap, rgb.width, rgb.height))]
            for half in halves:
                half.thumbnail((1150, 1800))
            gap = 12
            width = sum(half.width for half in halves) + gap
            height = max(half.height for half in halves)
            combined = Image.new("RGB", (width, height), "white")
            left = 0
            for label, half in zip(("1", "2"), halves):
                combined.paste(half, (left, 0))
                try:
                    from PIL import ImageDraw, ImageFont
                    font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 26)
                except Exception:
                    from PIL import ImageDraw, ImageFont
                    font = ImageFont.load_default()
                ImageDraw.Draw(combined).rectangle((left + 4, 4, left + 40, 38), fill="white")
                ImageDraw.Draw(combined).text((left + 10, 6), label, fill="black", font=font)
                left += half.width + gap
            rgb = combined
            anchor += "（上下两段，左段在上）"
        rgb.thumbnail((2300, 1800))
        output = BytesIO()
        rgb.save(output, format="JPEG", quality=86)
    return anchor, output.getvalue()


def bulk_word_rows(path, prefix):
    """Walk Word blocks through python-docx so merged cells repeat their value."""
    table = paragraph = 0
    for block in Document(path).iter_inner_content():
        if hasattr(block, "rows"):
            table += 1
            for index, row in enumerate(block.rows, 1):
                yield f"{prefix} / 表 {table} / 第 {index} 行", [cell.text for cell in row.cells]
        else:
            paragraph += 1
            yield f"{prefix} / 段落 {paragraph}", block.text


def bulk_text_chunks(path, unit, prefix, *, batch_rows=None, include_merges=False):
    """Walk a worksheet or Word body once and yield bounded, ordered chunks."""
    def chunks(iterator):
        start = 1
        while part := list(islice(iterator, batch_rows or BATCH)):
            yield start, part
            start += len(part)

    if path.suffix == ".doc":
        with converted_doc(path) as converted:
            check_office(converted, bulk=True)
            yield from bulk_text_chunks(converted, unit, prefix, batch_rows=batch_rows,
                                        include_merges=include_merges)
        return

    if unit["kind"] == "word":
        rows = wps_rows(path, prefix) if path.suffix == ".wps" else bulk_word_rows(path, prefix)
        yield from chunks(rows)
        return
    if path.suffix in LEGACY_SHEETS:
        yield from chunks(legacy_sheet_rows(path, unit["index"], prefix, 1, unit["size"],
                                            include_merges=include_merges))
        return
    book = load_workbook(path, read_only=not include_merges, data_only=True, keep_links=False)
    try:
        sheet = book.worksheets[unit["index"] - 1]
        merged = {}
        if include_merges:
            for area in sheet.merged_cells.ranges:
                if area.min_col != area.max_col:
                    continue
                value = value_text(sheet.cell(area.min_row, area.min_col))
                for row_number in range(area.min_row + 1, area.max_row + 1):
                    merged.setdefault(row_number, {})[area.min_col - 1] = value

        def rows():
            for i, row in enumerate(sheet.iter_rows(), 1):
                values = [value_text(cell) for cell in row]
                if any(values):
                    for col, value in merged.get(i, {}).items():
                        values[col] = value
                yield f"{prefix} / {sheet.title} / 第 {i} 行", values
        yield from chunks(rows())
    finally:
        book.close()
