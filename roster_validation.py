"""Shape and source validation for a roster submitted by the agent."""
from __future__ import annotations

import re
from collections.abc import Iterator
from typing import Any

from presets import PRESETS

MAX_RECORDS = 20_000


class RosterError(ValueError):
    """A submission that the model can repair using its tool result."""


CONTAINER_KEYS = {"sections", "groups", "data", "rows", "teams", "players", "pairs", "staff"}


def walk_records(value: Any) -> Iterator[dict]:
    if isinstance(value, list):
        for item in value:
            yield from walk_records(item)
    elif isinstance(value, dict):
        has_label = any(isinstance(value.get(field), str) and value[field].strip()
                        for field in ("name", "team", "contact", "partner", "source"))
        # A section may carry a display name ("部门组") plus its rows.  It is a
        # container, not a person; validating it as a record only creates false
        # "missing source" errors.
        is_container = any(key in value for key in CONTAINER_KEYS) and "source" not in value
        if has_label and not is_container:
            yield value
        for item in value.values():
            yield from walk_records(item)


def root_key(preset: str) -> str:
    return "groups" if preset == "two-column-numbered" else "sections"


def source_matches(source: str, anchors: dict) -> bool:
    """Accept exact anchors and a contiguous numeric range (e.g. 段落 4–5)."""
    if source in anchors:
        return True
    match = re.fullmatch(r"(.*?)(\d+)\s*[–—-]\s*(\d+)(.*)", source)
    if not match:
        return False
    prefix, first, last, suffix = match.groups()
    start, end = int(first), int(last)
    if start > end or end - start > 49:
        return False
    return all(f"{prefix}{number}{suffix}" in anchors for number in range(start, end + 1))


def _text(value: Any) -> str:
    return value.strip() if isinstance(value, str) else ""


def validate_content(value: Any, preset: str, path: str = "roster") -> None:
    """Reject structurally valid but unusable people/team records early."""
    if isinstance(value, list):
        for index, item in enumerate(value, 1):
            validate_content(item, preset, f"{path}[{index}]")
        return
    if not isinstance(value, dict):
        return
    if "players" in value:
        players = value.get("players")
        if not isinstance(players, list):
            raise RosterError(f"{path}.players 必须是数组。")
        for index, player in enumerate(players, 1):
            if not isinstance(player, dict) or not _text(player.get("name")):
                raise RosterError(f"{path}.players[{index}] 缺少姓名 name；不能只提交号码。")
    if "staff" in value:
        staff = value.get("staff")
        if not isinstance(staff, list):
            raise RosterError(f"{path}.staff 必须是数组。")
        for index, person in enumerate(staff, 1):
            if not isinstance(person, dict) or not _text(person.get("name")):
                raise RosterError(f"{path}.staff[{index}] 缺少姓名 name。")
    if "teams" in value:
        teams = value.get("teams")
        if not isinstance(teams, list):
            raise RosterError(f"{path}.teams 必须是数组。")
        for index, team in enumerate(teams, 1):
            field = "name" if preset == "two-column-numbered" else "team"
            if not isinstance(team, dict) or not (_text(team.get(field)) or
                    (field == "team" and _text(team.get("name")))):
                raise RosterError(f"{path}.teams[{index}] 缺少队名 {field}。")
    row_key = "pairs" if preset == "paired-table" and "pairs" in value else "rows"
    if row_key in value and preset in {"individual-table", "paired-table", "unit-contact-table"}:
        rows = value[row_key]
        if not isinstance(rows, list):
            raise RosterError(f"{path}.{row_key} 必须是数组。")
        for index, row in enumerate(rows, 1):
            if not isinstance(row, dict):
                raise RosterError(f"{path}.{row_key}[{index}] 不是对象。")
            if preset == "unit-contact-table":
                if not _text(row.get("team")) or not _text(row.get("contact")):
                    raise RosterError(f"{path}.{row_key}[{index}] 缺少单位 team 或联系人 contact。")
            elif not _text(row.get("name")):
                raise RosterError(f"{path}.{row_key}[{index}] 缺少姓名 name。")
    for key, child in value.items():
        if isinstance(child, (dict, list)):
            validate_content(child, preset, f"{path}.{key}")


def validate(value: Any, preset: str, anchors: dict) -> dict:
    if not isinstance(value, dict):
        raise RosterError("roster 必须是 JSON 对象。")
    root = root_key(preset)
    if set(value) != {root} or not isinstance(value[root], list) or not value[root]:
        raise RosterError(f"{preset} 母版要求顶层只有一个非空数组 {root}。")
    validate_content(value, preset)
    if preset == "two-column-numbered":
        teams_with_players: dict[str, bool] = {}
        for group in value[root]:
            if not isinstance(group, dict) or not group.get("teams"):
                raise RosterError("双栏编号每个分组必须包含非空 teams 数组。")
            for team in group["teams"]:
                name = _text(team["name"])
                teams_with_players[name] = teams_with_players.get(name, False) or bool(team.get("players"))
        missing = [name for name, has_players in teams_with_players.items() if not has_players]
        if missing:
            names = "、".join(missing[:5])
            raise RosterError(
                f"{names} 已提交队伍信息，但还未提交运动员。"
                "请将姓名和来源放在 entry.teams[].players 中，可用同一队名分批追加，再调用 finish_roster。"
            )

    records = list(walk_records(value))
    if not records:
        raise RosterError("名单里没有可识别的队伍或人员记录。")
    if len(records) > MAX_RECORDS:
        raise RosterError(f"记录数 {len(records)} 超过上限 {MAX_RECORDS}，请核对是否重复读取。")

    examples = list(anchors)[:8]
    hint = "、".join(examples) if examples else "none"
    for index, record in enumerate(records, 1):
        source = record.get("source")
        if not isinstance(source, str) or not source.strip():
            name = record.get("name") or record.get("team") or record.get("contact") or ""
            raise RosterError(
                f"第 {index} 条记录（{str(name)[:40]}）缺少 source。"
                f"请原样使用读取工具返回的 source；例如：{hint}。"
            )
        if not source_matches(source, anchors):
            raise RosterError(
                f"第 {index} 条记录的 source 不是本任务读取工具返回的编号：{source[:120]}。"
                f"请先读取包含该记录的页、表或行，再原样引用；连续范围可写成「段落 4–5」。可用示例：{hint}。"
            )
    return {"records": len(records), "root": root, "preset": PRESETS[preset][0]}
