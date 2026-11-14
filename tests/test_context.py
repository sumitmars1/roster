import json
import tempfile
import unittest
from pathlib import Path

import harness_runner
from limits import CONTEXT_WINDOW


class ContextTests(unittest.TestCase):
    def test_dsh_patch_uses_the_model_context_window(self):
        with tempfile.TemporaryDirectory() as directory:
            patch_path = harness_runner._patch(Path(directory))
            patch = json.loads(Path(patch_path).read_text(encoding="utf-8"))
        model = next(
            entry["config"]["providers"]["roster"]["models"][0]
            for item in patch for entry in item.get("insert", [])
            if entry.get("id") == "llm-pi-ai"
        )
        self.assertGreaterEqual(CONTEXT_WINDOW, 1_000_000)
        self.assertEqual(model["contextWindow"], CONTEXT_WINDOW)

    def test_dsh_patch_enables_native_retries_for_malformed_responses(self):
        with tempfile.TemporaryDirectory() as directory:
            patch_path = harness_runner._patch(Path(directory))
            patch = json.loads(Path(patch_path).read_text(encoding="utf-8"))
        provider = next(
            entry["config"]["providers"]["roster"]
            for item in patch for entry in item.get("insert", [])
            if entry.get("id") == "llm-pi-ai"
        )
        policy = provider["retryPolicy"]
        self.assertEqual(policy["mode"], "normal")
        self.assertEqual(policy["maxRetries"], 2)
        self.assertLessEqual(
            {"EMPTY_RESPONSE", "RATE_LIMIT", "SERVER", "TIMEOUT", "TRANSPORT", "MALFORMED_RESPONSE"},
            set(policy["retryableCodes"]),
        )
