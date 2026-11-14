import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import harness_runner
import mcp_server
import storage
from roster_validation import RosterError


class MultiMergeTests(unittest.TestCase):
    def test_staff_only_fragment_stays_open_until_players_are_appended(self):
        with (
            tempfile.TemporaryDirectory() as directory,
            patch.object(mcp_server, "SCOPE_FILES", ("F1",)),
        ):
            root = Path(directory)
            with patch.object(storage, "TASKS", root):
                folder = root / "task"
                source = "F1 报名表.jpg / 图片"
                storage.write_json(folder / "state.json", {"preset": "two-column-numbered"})
                storage.write_json(folder / "anchors.json", {source: {"file": "F1"}})
                mcp_server._append_roster("task", {"entry": {"teams": [
                    {"name": "甲队", "coach": "教练", "source": source}
                ]}})
                with self.assertRaisesRegex(RosterError, "还未提交运动员"):
                    mcp_server._finish_parts("task")
                self.assertFalse((folder / "fragments" / "F1.json").exists())
                mcp_server._append_roster("task", {"entry": {"teams": [
                    {"name": "甲队", "players": [{"name": "张三", "source": source}]}
                ]}})
                mcp_server._finish_parts("task")
                self.assertTrue((folder / "fragments" / "F1.json").is_file())

    def test_merge_requires_all_files_but_allows_one_fragment_for_multiple_pages(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            task = {
                "preset": "individual-table",
                "files": [{"id": file_id, "status": "done"} for file_id in ("F1", "F2")],
            }
            roster = {"sections": [{"data": {"title": "甲组", "rows": [{"name": "张三"}]}}]}
            fragment = folder / "fragments" / "F1.json"
            storage.write_json(fragment, {"file": "F1", "roster": roster})
            with self.assertRaisesRegex(ValueError, "F2"):
                harness_runner._merge_file_rosters(task, folder)
            storage.write_json(fragment, {"file": "F1", "files": ["F1", "F2"], "roster": roster})
            self.assertEqual(harness_runner._merge_file_rosters(task, folder), roster)

    def test_distinct_teams_on_the_same_source_page_stay_separate(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            task = {"preset": "two-column-numbered", "files": [{"id": "F1", "status": "done"}]}
            teams = [
                {"name": name, "source": "F1 报名表1.jpg / 图片", "players": []}
                for name in ("甲队", "乙队")
            ]
            storage.write_json(folder / "fragments" / "F1.json", {
                "file": "F1", "roster": {"groups": [{"teams": teams}]},
            })
            self.assertEqual(harness_runner._merge_file_rosters(task, folder)["groups"][0]["teams"], teams)

    def test_retry_preserves_review_instead_of_merging(self):
        task = {"files": [{"id": "F1", "status": "error"}, {"id": "F2", "status": "done"}]}
        with (
            patch.dict(harness_runner.os.environ, {"ROSTER_SINGLE_SESSION": "0"}),
            patch.object(storage, "read_task", return_value=task),
            patch.object(harness_runner, "_run_file", return_value="review"),
            patch.object(harness_runner, "_merge_and_render") as merge,
        ):
            harness_runner.retry_file("task", "F1")
            merge.assert_not_called()

    def test_same_team_from_two_fragments_does_not_duplicate_players(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            old = storage.TASKS
            storage.TASKS = root
            try:
                task_id = "task-1"
                folder = root / task_id
                (folder / "fragments").mkdir(parents=True)
                state = {
                    "id": task_id, "preset": "two-column-numbered", "title": "赛事", "group": "小学组",
                    "files": [
                        {"id": "F1", "status": "done"}, {"id": "F2", "status": "done"},
                    ],
                }
                (folder / "state.json").write_text(json.dumps(state), encoding="utf-8")
                roster = {"groups": [{"title": "小学组", "teams": [{
                    "name": "甲队", "source": "F1 第 1 页",
                    "players": [{"name": "张三", "number": "1", "source": "F1 第 1 页"}],
                }]}]}
                for file_id in ("F1", "F2"):
                    (folder / "fragments" / f"{file_id}.json").write_text(json.dumps({
                        "file": file_id, "roster": roster,
                    }, ensure_ascii=False), encoding="utf-8")
                merged = harness_runner._merge_file_rosters(state, folder)
                team = merged["groups"][0]["teams"][0]
                self.assertEqual(len(team["players"]), 1)
            finally:
                storage.TASKS = old

    def test_numbered_scan_files_merge_into_one_team(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            old = storage.TASKS
            storage.TASKS = root
            try:
                task_id = "task-2"
                folder = root / task_id
                (folder / "fragments").mkdir(parents=True)
                state = {
                    "id": task_id, "preset": "two-column-numbered", "title": "赛事", "group": "小学组",
                    "files": [{"id": "F1", "status": "done"}, {"id": "F2", "status": "done"}],
                }
                (folder / "state.json").write_text(json.dumps(state), encoding="utf-8")
                for index, name in enumerate(("江姐中学女子足球队", "自贡市江姐中学校"), 1):
                    roster = {"groups": [{"title": name, "teams": [{
                        "name": name, "source": f"F{index} 自贡市江姐中学校报名表{index}.jpg / 图片",
                        "players": [{"name": f"队员{index}", "number": str(index),
                                    "source": f"F{index} 报名表{index}.jpg / 图片"}],
                    }]}]}
                    (folder / "fragments" / f"F{index}.json").write_text(json.dumps({
                        "file": f"F{index}", "roster": roster,
                    }, ensure_ascii=False), encoding="utf-8")
                merged = harness_runner._merge_file_rosters(state, folder)
                teams = merged["groups"][0]["teams"]
                self.assertEqual(len(teams), 1)
                self.assertEqual(len(teams[0]["players"]), 2)
            finally:
                storage.TASKS = old

    def test_numbered_continuation_files_form_one_group(self):
        task = {"files": [
            {"id": "F10", "name": "自贡市江姐中学校报名表1.jpg"},
            {"id": "F11", "name": "自贡市江姐中学校报名表2.jpg"},
            {"id": "F12", "name": "自贡市江姐中学校报名表3.jpg"},
            {"id": "F1", "name": "成都市全兴小学校.pdf"},
        ]}
        group = harness_runner._file_group(task, "F11")
        self.assertEqual([item["id"] for item in group], ["F10", "F11", "F12"])
        self.assertEqual([item["id"] for item in harness_runner._file_group(task, "F1")], ["F1"])
