import unittest

from roster_validation import RosterError, source_matches, validate, walk_records


class ValidationTests(unittest.TestCase):
    def setUp(self):
        self.anchors = {
            "F1 报名表 / 段落 4": {},
            "F1 报名表 / 段落 5": {},
            "F1 报名表 / 第 2 页": {},
        }

    def test_section_name_is_not_a_person(self):
        value = {"sections": [{"name": "部门组", "data": {"title": "部门组", "rows": []}}]}
        self.assertEqual(list(walk_records(value)), [])

    def test_exact_source_and_contiguous_range(self):
        self.assertTrue(source_matches("F1 报名表 / 段落 4", self.anchors))
        self.assertTrue(source_matches("F1 报名表 / 段落 4-5", self.anchors))
        self.assertTrue(source_matches("F1 报名表 / 段落 4–5", self.anchors))
        self.assertFalse(source_matches("F1 报名表 / 段落 4-6", self.anchors))

    def test_validate_reports_missing_source(self):
        value = {"sections": [{"data": {"title": "组", "rows": [
            {"team": "甲队", "name": "张三", "source": "F1 报名表 / 段落 4"},
            {"team": "甲队", "name": "李四"},
        ]}}]}
        with self.assertRaisesRegex(RosterError, "缺少 source"):
            validate(value, "individual-table", self.anchors)

    def test_validate_counts_records(self):
        value = {"sections": [{"data": {"title": "组", "rows": [
            {"team": "甲队", "name": "张三", "source": "F1 报名表 / 段落 4"},
        ]}}]}
        result = validate(value, "individual-table", self.anchors)
        self.assertEqual(result["records"], 1)

    def test_player_needs_a_name(self):
        value = {"groups": [{"title": "组", "teams": [
            {"name": "甲队", "source": "F1 报名表 / 第 2 页", "players": [
                {"name": "张三", "number": "1", "source": "F1 报名表 / 第 2 页"},
                {"number": "2", "source": "F1 报名表 / 第 2 页"},
            ]},
        ]}]}
        with self.assertRaisesRegex(RosterError, "缺少姓名"):
            validate(value, "two-column-numbered", self.anchors)
