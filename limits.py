"""Model, tool-output and batch limits for the roster pipeline.

Defaults follow qwen3.8-flash on Bailian: 131072 max output tokens and a 1M
context window. The ceiling below leaves room for a whole small workbook in one
pass while staying far from the model limit.

Override with ROSTER_MAX_OUTPUT_TOKENS, ROSTER_MAX_TOOL_BYTES and
ROSTER_BATCH_ROWS.
"""
import os

DEFAULTS = {
    "ROSTER_MAX_OUTPUT_TOKENS": 32768,
    "ROSTER_MAX_TOOL_BYTES": 262144,
    "ROSTER_BATCH_ROWS": 200,
    "ROSTER_CONTEXT_WINDOW": 1_000_000,
}


def limit(name):
    default = DEFAULTS[name]
    try:
        value = int(os.getenv(name) or default)
    except ValueError:
        return default
    return value if value >= 1 else default


MAX_OUTPUT_TOKENS = limit("ROSTER_MAX_OUTPUT_TOKENS")
CONTEXT_WINDOW = limit("ROSTER_CONTEXT_WINDOW")
MAX_TOOL_BYTES = limit("ROSTER_MAX_TOOL_BYTES")
BATCH_ROWS = limit("ROSTER_BATCH_ROWS")
MODEL_BATCH_ROWS = BATCH_ROWS if os.getenv("ROSTER_BATCH_ROWS") else 40
