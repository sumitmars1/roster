"""Render the submitted roster with the bundled event-roster-docx skill."""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

from config import ROSTER_SKILL


class RenderError(ValueError):
    pass


def _apply_defaults(roster: dict, title: str, group: str) -> None:
    """Fill only missing headings; never rewrite a value the model supplied."""
    for section in roster.get("sections", []):
        data = section.get("data") if isinstance(section, dict) else None
        if isinstance(data, dict) and "title" in data and not str(data.get("title") or "").strip():
            data["title"] = title
    for item in roster.get("groups", []):
        if not isinstance(item, dict):
            continue
        if not str(item.get("title") or "").strip():
            raw_g = str(item.get("group") or item.get("name") or group or title or "代表队名单").strip()
            item["title"] = raw_g.split("；", 1)[0].strip() or "代表队名单"


def _pick(value: dict, allowed: set[str]) -> dict:
    return {key: value[key] for key in allowed if key in value}


def _players(value) -> list[dict]:
    return [_pick(item, {"name", "number", "source"}) for item in value or [] if isinstance(item, dict)]


def _staff(value) -> list[dict]:
    return [_pick(item, {"role", "name", "phone", "source"}) for item in value or [] if isinstance(item, dict)]


def _team(value: dict, allowed: set[str]) -> dict:
    team = _pick(value, allowed | {"team", "name", "source", "players", "staff"})
    if not team.get("team") and team.get("name"):
        team["team"] = team["name"]
    if "players" in team:
        team["players"] = _players(team["players"])
    if "staff" in team:
        team["staff"] = _staff(team["staff"])
    return team


def _sanitize(roster: dict, preset: str) -> None:
    """Drop fields the chosen DOCX template has no place for.

    Models often carry ID-card, birth-date, or alternate number fields that are
    useful evidence but not supported by a particular template. Failing the
    whole roster on that metadata loses otherwise valid data; strip it instead.
    """
    row_presets = {"individual-table", "paired-table", "unit-contact-table"}
    for section in roster.get("sections", []):
        if not isinstance(section, dict) or not isinstance(section.get("data"), dict):
            continue
        for key in list(section):
            if key not in {"data", "layout"}:
                section.pop(key, None)
        data = section["data"]
        row_key = "pairs" if preset == "paired-table" else "rows"
        if preset == "paired-table" and "pairs" not in data and isinstance(data.get("rows"), list):
            data["pairs"] = data.pop("rows")
        if preset in row_presets and isinstance(data.get(row_key), list):
            allowed = {"individual-table": {"team", "name", "event", "sign", "source"},
                       "paired-table": {"team", "name", "partner", "event", "sign", "source"},
                       "unit-contact-table": {"team", "contact", "phone", "source"}}[preset]
            rows = []
            for row in data[row_key]:
                if not isinstance(row, dict):
                    continue
                if preset != "unit-contact-table" and not row.get("sign") and row.get("number"):
                    row = {**row, "sign": row["number"]}
                rows.append(_pick(row, allowed))
            data[row_key] = rows
        if preset == "inline-roster" and isinstance(data.get("groups"), list):
            data["groups"] = [
                {"group": group.get("group") or group.get("title") or "", "teams": [
                    _team(team, {"team"}) for team in group.get("teams", []) if isinstance(team, dict)
                ]}
                for group in data["groups"] if isinstance(group, dict)
            ]
        if preset == "team-wide" and isinstance(data.get("teams"), list):
            data["teams"] = [_team(team, {"team", "leader", "phone", "coach", "coach_phone"})
                             for team in data["teams"] if isinstance(team, dict)]
        if preset == "grouped-team-table" and isinstance(data.get("groups"), list):
            data["groups"] = [
                {"group": group.get("group") or group.get("title") or "", "teams": [
                    _team(team, {"team"}) for team in group.get("teams", []) if isinstance(team, dict)
                ]}
                for group in data["groups"] if isinstance(group, dict)
            ]
    if preset == "two-column-numbered" and isinstance(roster.get("groups"), list):
        roster["groups"] = [
            {
                "title": str(group.get("title") or group.get("group") or group.get("name") or "").strip(),
                "teams": [
                    _team(team, {"name", "leader", "coach", "leader_phone", "coach_phone"})
                    for team in group.get("teams", []) if isinstance(team, dict)
                ],
            }
            for group in roster["groups"] if isinstance(group, dict)
        ]


def render(task_id: str, directory: Path, preset: str, roster: dict, title: str, group: str) -> dict:
    _apply_defaults(roster, title, group)
    _sanitize(roster, preset)
    source = directory / "roster.json"
    output = directory / "roster.docx"
    audit = directory / "placement.json"
    source.write_text(json.dumps(roster, ensure_ascii=False, indent=2), encoding="utf-8")
    for stale in (output, audit):
        stale.unlink(missing_ok=True)
    command = [
        sys.executable, str(ROSTER_SKILL / "scripts" / "roster.py"), "pages",
        "--preset", preset, "--data", str(source), "--output", str(output), "--audit", str(audit),
    ]
    try:
        result = subprocess.run(
            command, cwd=directory, capture_output=True, text=True, encoding="utf-8", timeout=120,
        )
    except subprocess.TimeoutExpired as exc:
        raise RenderError("Word 生成超时。") from exc
    if result.returncode != 0 or not output.is_file() or not audit.is_file():
        detail = (result.stderr or result.stdout or "生成工具没有返回 Word 文件").strip()
        detail = detail.replace(str(directory) + "/", "")[-1600:]
        raise RenderError(detail)
    try:
        report = json.loads(audit.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        report = {}
    shutil.rmtree(directory / "preview", ignore_errors=True)
    return {
        "word": output.name,
        "placement": audit.name,
        "pages": len(report.get("sections", [])) or len(report.get("groups", [])),
    }
