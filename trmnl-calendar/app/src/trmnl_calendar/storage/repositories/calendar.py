"""Read access to the calendar row and its categories (editing arrives in slice 2)."""

from __future__ import annotations

import json
import sqlite3
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from trmnl_calendar.domain.models import CalendarInfo, Section
from trmnl_calendar.errors import ValidationFailed

CALENDAR_ID = 1  # v1 has one calendar (A-1); tables still carry calendar_id


def get_calendar(conn: sqlite3.Connection) -> CalendarInfo:
    row = conn.execute("SELECT * FROM calendars WHERE id = ?", (CALENDAR_ID,)).fetchone()
    opts = json.loads(row["options_json"] or "{}")
    week_start = opts.get("week_start", "sun")
    time_format = opts.get("time_format", "12")
    return CalendarInfo(
        id=row["id"],
        name=row["name"],
        timezone=row["timezone"],
        week_start=week_start if week_start in ("sun", "mon") else "sun",
        time_format=time_format if time_format in ("12", "24") else "12",
    )


def list_sections(conn: sqlite3.Connection, *, include_hidden: bool = True) -> list[Section]:
    sql = "SELECT * FROM sections WHERE calendar_id = ?"
    if not include_hidden:
        sql += " AND visible = 1"
    rows = conn.execute(sql + " ORDER BY position, id", (CALENDAR_ID,)).fetchall()
    return [
        Section(r["id"], r["name"], r["position"], r["weight"], bool(r["visible"]), r["placement"])
        for r in rows
    ]


def update_calendar(conn: sqlite3.Connection, patch: dict) -> CalendarInfo:
    cur = get_calendar(conn)
    unknown = set(patch) - {"name", "timezone", "week_start", "time_format"}
    if unknown:
        raise ValidationFailed({k: "Unknown field" for k in sorted(unknown)})
    err: dict[str, str] = {}
    name = patch.get("name", cur.name)
    name = name.strip() if isinstance(name, str) else ""
    if not 1 <= len(name) <= 40:
        err["name"] = "Enter a name up to 40 characters"
    tz = patch.get("timezone", cur.timezone)
    tz = tz.strip() if isinstance(tz, str) else ""
    try:
        ZoneInfo(tz)
    except (ZoneInfoNotFoundError, ValueError, OSError):
        err["timezone"] = "Unknown timezone. Use a name like America/Halifax"
    week_start = patch.get("week_start", cur.week_start)
    if week_start not in ("sun", "mon"):
        err["week_start"] = "Choose Sunday or Monday"
    time_format = patch.get("time_format", cur.time_format)
    if time_format not in ("12", "24"):
        err["time_format"] = "Choose 12-hour or 24-hour"
    if err:
        raise ValidationFailed(err)
    row = conn.execute("SELECT options_json FROM calendars WHERE id = ?", (CALENDAR_ID,)).fetchone()
    opts = json.loads(row["options_json"] or "{}")
    opts.update(week_start=week_start, time_format=time_format)
    conn.execute(
        "UPDATE calendars SET name = ?, timezone = ?, options_json = ? WHERE id = ?",
        (name, tz, json.dumps(opts), CALENDAR_ID),
    )
    conn.commit()
    return get_calendar(conn)
