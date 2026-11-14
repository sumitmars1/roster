from __future__ import annotations

import asyncio
import json
import os
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pymupdf
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

import book_runner
import config
import documents
import harness_runner
import mcp_server
import renderer
import storage


class ParallelAndShardingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.tasks_dir = self.root / "tasks"
        self.tasks_dir.mkdir(parents=True, exist_ok=True)
        self.orig_config_data = config.DATA
        self.orig_config_tasks = config.TASKS
        self.orig_storage_tasks = storage.TASKS
        self.orig_harness_data = harness_runner.DATA
        config.DATA = self.root
        config.TASKS = self.tasks_dir
        storage.TASKS = self.tasks_dir
        harness_runner.DATA = self.root

    def tearDown(self) -> None:
        config.DATA = self.orig_config_data
        config.TASKS = self.orig_config_tasks
        storage.TASKS = self.orig_storage_tasks
        harness_runner.DATA = self.orig_harness_data
        self.tmp.cleanup()

    def _make_task_dir(self, task_id: str) -> Path:
        tdir = self.tasks_dir / task_id
        tdir.mkdir(parents=True, exist_ok=True)
        return tdir

    def _build_merged_workbook(self, path: Path) -> Path:
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "街镇"
        ws1.append(["序号", "代表队", "姓名", "号码"])
        row_idx = 2
        for team_no in range(1, 25):
            team_name = f"街镇第{team_no}队"
            start_r = row_idx
            members = 6 if team_no % 2 == 1 else 5
            for m in range(members):
                ws1.append([team_no if m == 0 else None, team_name if m == 0 else None, f"张三{team_no}_{m}", f"{(team_no * 10 + m) % 100:02d}"])
                ws1.cell(row=row_idx, column=4).number_format = "@"
                row_idx += 1
            ws1.merge_cells(start_row=start_r, start_column=1, end_row=row_idx - 1, end_column=1)
            ws1.merge_cells(start_row=start_r, start_column=2, end_row=row_idx - 1, end_column=2)

        ws2 = wb.create_sheet("部门")
        ws2.append(["序号", "代表队", "姓名", "号码"])
        ws2.append(["男子甲组", None, None, None])
        row_idx2 = 3
        for team_no in range(1, 31):
            if team_no == 16:
                ws2.append(["男子乙组", None, None, None])
                row_idx2 += 1
            team_name = f"部门第{team_no}队"
            start_r = row_idx2
            for m in range(7):
                ws2.append([team_no if m == 0 else None, team_name if m == 0 else None, f"李四{team_no}_{m}", f"{team_no:02d}{m}"])
                ws2.cell(row=row_idx2, column=4).number_format = "@"
                row_idx2 += 1
            ws2.merge_cells(start_row=start_r, start_column=1, end_row=row_idx2 - 1, end_column=1)
            ws2.merge_cells(start_row=start_r, start_column=2, end_row=row_idx2 - 1, end_column=2)

        ws3 = wb.create_sheet("候补未参赛")
        ws3.append(["序号", "代表队", "姓名"])
        ws3.append([1, "备选队", "王五"])
        wb.save(path)
        wb.close()
        return path

    def test_storage_locked_and_append_trace_thread_safe(self) -> None:
        tdir = self._make_task_dir("concurrent_t1")
        storage.write_json(tdir / "state.json", {"id": "concurrent_t1", "files": [{"id": "F1", "status": "waiting"}, {"id": "F2", "status": "waiting"}]})

        errors: list[Exception] = []

        def worker(idx: int) -> None:
            try:
                for step in range(25):
                    storage.append_trace("concurrent_t1", {"kind": "tool", "worker": idx, "step": step})
                    harness_runner._update_file("concurrent_t1", "F1" if idx % 2 == 0 else "F2", status=f"s{step}")
            except Exception as exc:
                errors.append(exc)

        threads = [threading.Thread(target=worker, args=(i,)) for i in range(10)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        self.assertEqual(errors, [])
        lines = (tdir / "trace.jsonl").read_text(encoding="utf-8").splitlines()
        self.assertEqual(len(lines), 250)
        for line in lines:
            parsed = json.loads(line)
            self.assertIn("worker", parsed)

    def test_excel_sharding_preserves_all_rows_merges_and_leading_zeros(self) -> None:
        xlsx = self._build_merged_workbook(self.root / "reg.xlsx")
        out_dir = self.root / "shards"
        out_dir.mkdir()

        shards = book_runner._prepare_excel_shards(
            xlsx,
            "报名汇总表.xlsx",
            included_sheets=["街镇", "部门"],
            excluded_sheets=["候补未参赛"],
            norm_gmap={"街镇": "街镇男子组", "部门": "部门男子组"},
            fallback_group="默认组",
            out_dir=out_dir,
            file_prefix="src1",
        )
        self.assertIsNotNone(shards)
        assert shards is not None

        street_shards = [s for s in shards if s["group"] == "街镇男子组"]
        dept_shards = [s for s in shards if s["group"] == "部门男子组"]
        self.assertGreaterEqual(len(street_shards), 2)
        self.assertGreaterEqual(len(dept_shards), 3)
        self.assertTrue(all("男子组" in s["context_hint"] for s in shards))
        self.assertTrue(any("男子乙组" in s["context_hint"] or "男子甲组" in s["context_hint"] for s in dept_shards))

        extracted_players_street: list[tuple[str, str, str]] = []
        for s in street_shards:
            wb = openpyxl.load_workbook(s["path"], data_only=True)
            ws = wb.active
            self.assertLessEqual(ws.max_row, 68)
            wb.close()
            rows, nxt, _, _ = documents.extract(s["path"], s["units"][0], 1, "F1", bulk=True, row_limit=150)
            self.assertIsNone(nxt)
            self.assertEqual(rows[0][1][:4], ["序号", "代表队", "姓名", "号码"])
            for _, vals in rows[1:]:
                if len(vals) >= 4 and vals[2]:
                    extracted_players_street.append((vals[1], vals[2], vals[3]))

        wb_orig = openpyxl.load_workbook(xlsx, data_only=True)
        ws_orig = wb_orig["街镇"]
        orig_count = sum(1 for r in range(2, ws_orig.max_row + 1) if ws_orig.cell(r, 3).value)
        wb_orig.close()
        self.assertEqual(len(extracted_players_street), orig_count)
        self.assertIn(("街镇第1队", "张三1_0", "10"), extracted_players_street)
        self.assertIn(("街镇第10队", "张三10_1", "01"), extracted_players_street)
        self.assertTrue(any(p[2].startswith("0") for p in extracted_players_street))

    def test_pdf_sharding_splits_large_pdfs_only(self) -> None:
        small_pdf = self.root / "small.pdf"
        doc_small = pymupdf.open()
        for i in range(6):
            p = doc_small.new_page()
            p.insert_text((72, 72), f"Page {i + 1}")
        doc_small.save(small_pdf)
        doc_small.close()

        large_pdf = self.root / "large.pdf"
        doc_large = pymupdf.open()
        for i in range(18):
            p = doc_large.new_page()
            p.insert_text((72, 72), f"BigPage {i + 1}")
        doc_large.save(large_pdf)
        doc_large.close()

        out_dir = self.root / "pdf_shards"
        out_dir.mkdir()
        self.assertIsNone(book_runner._prepare_pdf_shards(small_pdf, "small.pdf", "甲组", out_dir, "s1"))
        big_shards = book_runner._prepare_pdf_shards(large_pdf, "large.pdf", "甲组", out_dir, "s2")
        self.assertIsNotNone(big_shards)
        assert big_shards is not None
        self.assertEqual([sum(1 for _ in s["units"]) for s in big_shards], [8, 8, 2])

    def test_stage1_digest_and_book_mcp_anti_loop(self) -> None:
        xlsx = self._build_merged_workbook(self.root / "digest.xlsx")
        pdf_path = self.root / "rules.pdf"
        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text((72, 72), "Competition Rules Article 1")
        doc.save(pdf_path)
        doc.close()

        tdir = self._make_task_dir("bookplan_p1_v1")
        shutil_dest_x = tdir / "f01.xlsx"
        shutil_dest_x.write_bytes(xlsx.read_bytes())
        shutil_dest_p = tdir / "f02.pdf"
        shutil_dest_p.write_bytes(pdf_path.read_bytes())
        storage.write_json(
            tdir / "state.json",
            {
                "id": "bookplan_p1_v1",
                "files": [
                    {"id": "F1", "name": "报名表.xlsx", "stored_name": "f01.xlsx", "suffix": ".xlsx", "units": documents.inventory(shutil_dest_x, bulk=True)},
                    {"id": "F2", "name": "规程.pdf", "stored_name": "f02.pdf", "suffix": ".pdf", "units": documents.inventory(shutil_dest_p, bulk=True)},
                ],
                "template_spec": {"id": "tpl", "name": "通用模板", "roster_preset": "two-column-numbered", "chapters": []},
            },
        )
        digest = book_runner._build_stage1_digest(tdir, storage.read_task("bookplan_p1_v1")["files"])
        self.assertIn("工作表「街镇」", digest)
        self.assertIn("工作表「部门」", digest)
        self.assertIn("Competition Rules Article 1", digest)

        async def _check_mcp() -> None:
            params = StdioServerParameters(
                command=sys.executable,
                args=[str(Path(__file__).resolve().parents[1] / "book_mcp.py"), "--task", "bookplan_p1_v1", "--mode", "plan"],
                env=dict(os.environ, EVENT_BOOK_DATA_DIR=str(self.root)),
            )
            async with stdio_client(params) as (read, write):
                async with ClientSession(read, write) as sess:
                    await sess.initialize()
                    res1 = await sess.call_tool("read_document", {"file_id": "F1", "unit": "街镇", "start": 1, "limit": 5})
                    self.assertFalse(res1.is_error)
                    p1 = json.loads(res1.content[0].text)
                    self.assertGreaterEqual(len(p1["rows"]), 80)
                    self.assertNotIn("plan_hint", p1)

                    res2 = await sess.call_tool("read_document", {"file_id": "F1", "unit": "街镇", "start": 20})
                    self.assertFalse(res2.is_error)
                    p2 = json.loads(res2.content[0].text)
                    self.assertIn("plan_hint", p2)

                    for step in range(3, 6):
                        res_mid = await sess.call_tool("read_document", {"file_id": "F1", "unit": "街镇", "start": step * 10})
                        self.assertFalse(res_mid.is_error)

                    res6 = await sess.call_tool("read_document", {"file_id": "F1", "unit": "街镇", "start": 60})
                    self.assertTrue(res6.is_error)
                    self.assertIn("submit_book_plan", res6.content[0].text)

        asyncio.run(_check_mcp())

    def test_parallel_run_multi_and_stage2_multigroup_split_and_merged(self) -> None:
        self.assertEqual(harness_runner._file_group_key({"id": "F1", "name": "reg_p1.xlsx", "shard_id": "s1"}), "shard:F1")
        self.assertEqual(harness_runner._file_group_key({"id": "F2", "name": "reg_p2.xlsx", "shard_id": "s2"}), "shard:F2")

        tdir = self._make_task_dir("parallel_multi")
        storage.write_json(
            tdir / "state.json",
            {
                "id": "parallel_multi",
                "preset": "individual-table",
                "title": "测试赛",
                "group": "甲组",
                "status": "queued",
                "files": [
                    {"id": "F1", "name": "reg_p1.xlsx", "shard_id": "s1", "status": "waiting"},
                    {"id": "F2", "name": "reg_p2.xlsx", "shard_id": "s2", "status": "waiting"},
                    {"id": "F3", "name": "reg_p3.xlsx", "shard_id": "s3", "status": "waiting"},
                ],
            },
        )

        def fake_run_file(task_id: str, file_id: str, *, resume: bool = False, answer: str = "", clear_review: bool = True) -> str:
            time.sleep(0.2)
            storage.write_json(
                tdir / "fragments" / f"{file_id}.json",
                {
                    "file": file_id,
                    "files": [file_id],
                    "records": 1,
                    "roster": {
                        "sections": [
                            {"data": {"title": "男子甲组", "rows": [{"name": f"选手_{file_id}", "source": f"{file_id} / 第 1 页"}]}}
                        ]
                    },
                },
            )
            harness_runner._update_file(task_id, file_id, status="done")
            return "done"

        started = time.monotonic()
        with patch.object(harness_runner, "_run_file", side_effect=fake_run_file), patch.object(
            harness_runner, "validate", return_value={"records": 3}
        ), patch.object(harness_runner, "render", return_value={"word": "roster.docx"}):
            harness_runner._run_multi("parallel_multi")
        elapsed = time.monotonic() - started
        self.assertLess(elapsed, 0.45)

        merged = harness_runner._merge_file_rosters(storage.read_task("parallel_multi"), tdir)
        self.assertEqual(len(merged["sections"]), 1)
        self.assertEqual(len(merged["sections"][0]["data"]["rows"]), 3)

        # Verify Stage 2 runs separate competition groups in parallel and handles split_by_group_in_toc=True vs False
        xlsx = self._build_merged_workbook(self.root / "book_src.xlsx")
        plan = {
            "event_title": "区运会",
            "roster_scope": {
                "groups": ["街镇男子组", "部门男子组"],
                "sources": [
                    {
                        "file_id": "F1",
                        "included_sheets": ["街镇", "部门"],
                        "excluded_sheets": ["候补未参赛"],
                        "group_map": {" 街镇 ": "街镇男子组", "部门": "部门男子组"},
                    }
                ],
            },
        }
        project = {"id": "p1", "name": "区运会"}
        tpl_split = {
            "id": "t1",
            "roster_preset": "two-column-numbered",
            "chapters": [{"id": "roster", "kind": "roster", "split_by_group_in_toc": True}],
        }
        tpl_combined = {
            "id": "t2",
            "roster_preset": "two-column-numbered",
            "chapters": [{"id": "roster", "kind": "roster", "split_by_group_in_toc": False}],
        }

        def fake_run_task(tid: str) -> None:
            time.sleep(0.25)
            st = storage.read_task(tid)
            task_folder = config.TASKS / tid
            gname = st["group"]
            r_json = {"groups": [{"title": gname, "teams": [{"name": f"{gname}一队", "source": "F1 / 第 1 页", "players": [{"name": "选手A", "source": "F1 / 第 1 页"}]}]}]}
            storage.write_json(task_folder / "roster.json", r_json)
            (task_folder / "roster.docx").write_bytes(b"docx")
            storage.update_task(tid, status="done", result={"word": "roster.docx", "records": 1})

        t0 = time.monotonic()
        with patch.object(harness_runner, "run_task", side_effect=fake_run_task):
            res_split = book_runner._run_stage2_rosters(
                project, tpl_split, plan, {"F1": xlsx}, {"F1": "报名表.xlsx"}, "bookroster_p1_v1"
            )
        self.assertLess(time.monotonic() - t0, 0.45)
        self.assertEqual([r["group"] for r in res_split], ["街镇男子组", "部门男子组"])
        self.assertTrue(all(r["docx"].is_file() for r in res_split))

        def fake_render(task_id: str, directory: Path, preset: str, roster: dict, title: str, group: str) -> dict:
            out = directory / "roster.docx"
            out.write_bytes(b"combined")
            return {"word": "roster.docx"}

        with patch.object(harness_runner, "run_task", side_effect=fake_run_task), patch.object(
            renderer, "render", side_effect=fake_render
        ):
            res_combined = book_runner._run_stage2_rosters(
                project, tpl_combined, plan, {"F1": xlsx}, {"F1": "报名表.xlsx"}, "bookroster_p1_v2"
            )
        self.assertEqual(len(res_combined), 1)
        self.assertEqual(res_combined[0]["group"], "")
        combined_json = storage.read_json(config.TASKS / "bookroster_p1_v2" / "roster.json")
        self.assertEqual([g["title"] for g in combined_json["groups"]], ["街镇男子组", "部门男子组"])

    def test_mcp_server_batch_limits(self) -> None:
        self.assertEqual(mcp_server.MAX_CHUNK_BYTES, 64_000)
        self.assertEqual(mcp_server.MAX_APPEND_ROWS, 80)
        tdir = self._make_task_dir("batch_rows")
        storage.write_json(tdir / "state.json", {"id": "batch_rows", "preset": "individual-table", "files": []})
        storage.write_json(tdir / "anchors.json", {"F1 / 表 / 第 1 行": True})
        rows = [{"name": f"选手{i}", "source": "F1 / 表 / 第 1 行"} for i in range(60)]
        msg = mcp_server._append_rows("batch_rows", {"group": "甲组", "rows": rows})
        self.assertIn("累计 60 行", msg)


if __name__ == "__main__":
    unittest.main()
