import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document

import mcp_server
import storage
from presets import PRESETS
from renderer import render
from roster_validation import validate


class RendererTests(unittest.TestCase):
    def test_builtin_preset_examples_validate_and_render(self):
        for preset, (_, _, example) in PRESETS.items():
            with self.subTest(preset=preset), tempfile.TemporaryDirectory() as directory:
                roster = json.loads(example)
                self.assertGreater(validate(roster, preset, {"定位": {}})["records"], 0)
                folder = Path(directory)
                result = render("test", folder, preset, roster, "测试赛事", "甲组")
                document = Document(folder / result["word"])
                self.assertTrue(document.tables or any(p.text.strip() for p in document.paragraphs))

    def test_paired_template_renders_appended_batches_and_legacy_rows(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(storage, "TASKS", Path(directory)):
            folder = Path(directory) / "test"
            storage.write_json(folder / "state.json", {"preset": "paired-table"})
            storage.write_json(folder / "anchors.json", {"定位": {}})
            storage.write_json(folder / "parts.json", [{"data": {"title": "乙组", "rows": [
                {"team": "乙队", "name": "郑五", "partner": "王七", "source": "定位"}
            ]}}])
            mcp_server._append_rows("test", {"title": "甲组", "rows": [
                {"team": "甲队", "name": "张三", "partner": "李四", "source": "定位"}
            ]})
            mcp_server._append_rows("test", {"title": "甲组", "rows": [
                {"team": "甲队", "name": "田六", "partner": "周八", "source": "定位"}
            ]})
            mcp_server._finish_parts("test")
            document = Document(folder / "roster.docx")
            contents = "".join(cell.text for table in document.tables for row in table.rows for cell in row.cells)
            for name in ("张三", "李四", "郑五", "王七", "田六", "周八"):
                self.assertIn(name, contents.replace("\u3000", ""))

    def test_individual_table_renders(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            roster = {"sections": [{"data": {"title": "甲组", "rows": [
                {"team": "甲队", "name": "张三", "event": "单打", "sign": "001", "source": "F1 第 2 行"},
                {"team": "甲队", "name": "李四", "event": "单打", "sign": "002", "source": "F1 第 3 行"},
            ]}}]}
            result = render("test", folder, "individual-table", roster, "测试赛事", "甲组")
            self.assertTrue((folder / result["word"]).is_file())
            document = Document(folder / result["word"])
            text = "\n".join(p.text for p in document.paragraphs) + "\n" + "\n".join(
                cell.text for table in document.tables for row in table.rows for cell in row.cells
            )
            self.assertIn("张三", text.replace("\u3000", ""))
            self.assertIn("001", text)
