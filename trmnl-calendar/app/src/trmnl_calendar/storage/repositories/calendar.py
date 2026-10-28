"""Read access to the calendar row and its categories (editing arrives in slice 2)."""

from __future__ import annotations

import json
import sqlite3

from trmnl_calendar.domain.models import CalendarInfo, Section

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
