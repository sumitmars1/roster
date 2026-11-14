"""Runtime configuration for the Harness roster application."""
from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
TASKS = DATA / "tasks"
STATIC = ROOT / "static"
SKILLS = ROOT / "skills" / "intake"
ROSTER_SKILL = ROOT / "roster_skill"
MCP_SERVER = ROOT / "mcp_server.py"

DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DEFAULT_MODEL = "qwen3.8-flash"

MAX_FILES = 100
MAX_FILE_BYTES = 300 * 1024**2
MAX_TOTAL_BYTES = 3 * 1024**3
MAX_REVIEW_ANSWER = 2500


def load_env(path: Path) -> None:
    """Load a small KEY=VALUE file without exporting unrelated shell state."""
    if not path.is_file():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_env(ROOT / ".env")

API_KEY = os.getenv("DASHSCOPE_API_KEY", "").strip()
BASE_URL = os.getenv("ROSTER_BASE_URL", DEFAULT_BASE_URL).rstrip("/")
MODEL = os.getenv("ROSTER_MODEL", DEFAULT_MODEL)
REQUEST_TIMEOUT_SECONDS = 180.0


def ensure_dirs() -> None:
    TASKS.mkdir(parents=True, exist_ok=True)
