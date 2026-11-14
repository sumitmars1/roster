import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pymupdf
from docx import Document
from openpyxl import Workbook
from PIL import Image

import document_bank
import documents
import mcp_server
import storage
from roster_validation import RosterError


class DocumentBankTests(unittest.TestCase):
    def test_word_rows_and_anchors(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            old = storage.TASKS
            storage.TASKS = root
            try:
                task = root / "task-1"
                task.mkdir()
                source = task / "f01.docx"
                document = Document()
                table = document.add_table(rows=3, cols=2)
                table.cell(0, 0).text = "姓名"
                table.cell(0, 1).text = "号码"
                table.cell(1, 0).text = "张三"
                table.cell(1, 1).text = "07"
                table.cell(2, 0).text = "李四"
                table.cell(2, 1).text = "08"
                document.save(source)
                (task / "state.json").write_text(json.dumps({
                    "id": "task-1", "title": "测试", "group": "甲组", "preset": "inline-roster",
                    "files": [{"id": "F1", "name": "名单.docx", "stored_name": "f01.docx",
                               "suffix": ".docx", "size": source.stat().st_size,
                               "units": documents.inventory(source, bulk=True)}],
                }, ensure_ascii=False), encoding="utf-8")
                payload, image = document_bank.read_document("task-1", {"file_id": "F1", "unit": 1})
                self.assertIsNone(image)
                self.assertTrue(payload["rows"])
                self.assertTrue(any("张三" in str(row["text"]) for row in payload["rows"]))
                self.assertTrue(document_bank.anchors("task-1"))
                self.assertIn(1, document_bank.progress("task-1")["F1"])
            finally:
                storage.TASKS = old

    def test_office_ranges_must_cover_the_unit_before_submission(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(storage, "TASKS", Path(directory)):
            task = Path(directory) / "task"
            task.mkdir()
            source = task / "f01.docx"
            doc = Document()
            for name in ("张三", "李四", "王五"):
                doc.add_paragraph(name * 100)
            doc.save(source)
            storage.write_json(task / "state.json", {
                "preset": "individual-table",
                "files": [{"id": "F1", "name": "报名.docx", "stored_name": source.name,
                           "units": documents.inventory(source, bulk=True)}],
            })
            args = {"file_id": "F1", "unit": 1}
            with patch.object(document_bank, "MAX_TOOL_BYTES", 1300):
                first, _ = document_bank.read_document("task", args)
                self.assertEqual(first["next_start"], 2)
                self.assertEqual(len(document_bank.anchors("task")), 2)
                document_bank.read_document("task", {**args, "start": 3})
                self.assertEqual(document_bank.progress("task")["F1"], [])
                roster = {"sections": [{"data": {"title": "名单", "rows": [
                    {"name": "张三", "source": first["rows"][0]["source"]}
                ]}}]}
                with self.assertRaisesRegex(RosterError, "未读完"):
                    mcp_server._finish("task", roster)
                document_bank.read_document("task", {**args, "start": 2})
                self.assertEqual(document_bank.progress("task")["F1"], [1])

    def test_pdf_text_paginates_and_previous_page_returns_its_tail(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(storage, "TASKS", Path(directory)):
            task = Path(directory) / "task"
            task.mkdir()
            source = task / "f01.pdf"
            with pymupdf.open() as pdf:
                page = pdf.new_page()
                for i in range(1, 13):
                    page.insert_text((30, 20 * i), f"page-one-line-{i}")
                pdf.new_page().insert_text((30, 20), "page-two")
                pdf.save(source)
            storage.write_json(task / "state.json", {
                "files": [{"id": "F1", "name": "报名.pdf", "stored_name": source.name,
                           "units": documents.inventory(source, bulk=True)}],
            })
            args = {"file_id": "F1", "unit": 1, "limit": 5}
            first, _ = document_bank.read_document("task", args)
            self.assertEqual(first["next_start"], 6)
            self.assertEqual(document_bank.progress("task")["F1"], [])
            last, _ = document_bank.read_document("task", {**args, "start": 11})
            self.assertIsNone(last["next_start"])
            document_bank.read_document("task", {**args, "start": 6})
            self.assertEqual(document_bank.progress("task")["F1"], [1])
            second, _ = document_bank.read_document("task", {**args, "unit": 2})
            self.assertEqual(second["adjacent"]["before"][-1]["text"], "page-one-line-12")

    def test_excel_preserves_vertical_merge_and_formatted_number(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(storage, "TASKS", Path(directory)):
            task = Path(directory) / "task"
            task.mkdir()
            source = task / "f01.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.append(["甲队", "张三", 7])
            sheet.append([None, "李四", 8])
            sheet.merge_cells("A1:A2")
            sheet["C2"].number_format = "000"
            book.save(source)
            book.close()
            storage.write_json(task / "state.json", {
                "files": [{"id": "F1", "name": "报名.xlsx", "stored_name": source.name,
                           "units": documents.inventory(source, bulk=True)}],
            })
            result, _ = document_bank.read_document("task", {"file_id": "F1", "unit": 1, "start": 2})
            self.assertEqual(result["rows"][0]["text"], ["甲队", "李四", "008"])
            self.assertEqual(document_bank.progress("task")["F1"], [])

    def test_pdf_with_text_and_large_image_needs_visual_read(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(storage, "TASKS", Path(directory)):
            task = Path(directory) / "task"
            task.mkdir()
            image_path = task / "scan.png"
            Image.new("RGB", (50, 50), "white").save(image_path)
            source = task / "f01.pdf"
            with pymupdf.open() as pdf:
                page = pdf.new_page()
                page.insert_text((30, 20), "Roster")
                page.insert_image(pymupdf.Rect(30, 50, 560, 800), filename=image_path)
                pdf.save(source)
            storage.write_json(task / "state.json", {
                "files": [{"id": "F1", "name": "报名.pdf", "stored_name": source.name,
                           "units": documents.inventory(source, bulk=True)}],
            })
            args = {"file_id": "F1", "unit": 1}
            for view in ("text", "layout"):
                result, _ = document_bank.read_document("task", {**args, "view": view})
                self.assertTrue(result["needs_image"])
                self.assertEqual(document_bank.progress("task").get("F1", []), [])
            _, image = document_bank.read_document("task", {**args, "view": "image"})
            self.assertTrue(image)
            self.assertEqual(document_bank.progress("task")["F1"], [1])
