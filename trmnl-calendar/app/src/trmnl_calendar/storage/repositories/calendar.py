"""Read access to the calendar row and its categories (editing arrives in slice 2)."""

from __future__ import annotations

import json
import re
import sqlite3
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from trmnl_calendar.domain.models import CalendarInfo, Section
from trmnl_calendar.errors import ValidationFailed

CALENDAR_ID = 1  # v1 has one calendar (A-1); tables still carry calendar_id
FIELDS = ("name", "timezone", "week_start", "time_format", "home_airport", "flight_monthly_limit")
_AIRPORT = re.compile(r"[A-Z]{3,4}")
LIMIT_MAX = 1000


def _limit(value) -> int | None:
    try:
        n = int(str(value).strip())
    except (TypeError, ValueError):
        return None
    return n if 1 <= n <= LIMIT_MAX else None


def get_calendar(conn: sqlite3.Connection) -> CalendarInfo:
    row = conn.execute("SELECT * FROM calendars WHERE id = ?", (CALENDAR_ID,)).fetchone()
    opts = json.loads(row["options_json"] or "{}")
    week_start = opts.get("week_start", "sun")
    time_format = opts.get("time_format", "12")
    home = opts.get("home_airport", "")
    return CalendarInfo(
        id=row["id"],
        name=row["name"],
        timezone=row["timezone"],
        week_start=week_start if week_start in ("sun", "mon") else "sun",
        time_format=time_format if time_format in ("12", "24") else "12",
        home_airport=home if isinstance(home, str) and _AIRPORT.fullmatch(home) else "",
        flight_monthly_limit=_limit(opts.get("flight_monthly_limit", 100)) or 100,
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
    unknown = set(patch) - set(FIELDS)
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
    home = patch.get("home_airport", cur.home_airport)
    home = home.strip().upper() if isinstance(home, str) else ""
    if home and not _AIRPORT.fullmatch(home):
        err["home_airport"] = "Enter an airport code like YHZ or CYHZ, or leave it empty"
    limit = _limit(patch.get("flight_monthly_limit", cur.flight_monthly_limit))
    if limit is None:
        err["flight_monthly_limit"] = f"Enter a whole number from 1 to {LIMIT_MAX}"
    if err:
        raise ValidationFailed(err)
    row = conn.execute("SELECT options_json FROM calendars WHERE id = ?", (CALENDAR_ID,)).fetchone()
    opts = json.loads(row["options_json"] or "{}")
    opts.update(
        week_start=week_start,
        time_format=time_format,
        home_airport=home,
        flight_monthly_limit=limit,
    )
    conn.execute(
        "UPDATE calendars SET name = ?, timezone = ?, options_json = ? WHERE id = ?",
        (name, tz, json.dumps(opts), CALENDAR_ID),
    )
    conn.commit()
    return get_calendar(conn)


def seed_options(conn: sqlite3.Connection, defaults: dict) -> list[str]:
    """Save environment defaults for options never saved before; a saved value wins (ADR-0005)."""
    row = conn.execute("SELECT options_json FROM calendars WHERE id = ?", (CALENDAR_ID,)).fetchone()
    opts = json.loads(row["options_json"] or "{}")
    added = [k for k, v in defaults.items() if k not in opts and v not in (None, "")]
    if added:
        opts.update({k: defaults[k] for k in added})
        conn.execute(
            "UPDATE calendars SET options_json = ? WHERE id = ?", (json.dumps(opts), CALENDAR_ID)
        )
        conn.commit()
    return added
