"""Everything the owner typed, as plain JSON: a safety net and a way out (docs/06 `GET /export`).

Not included: push state, idempotency keys and app_meta (they are rebuilt by the app).
"""

from __future__ import annotations

import sqlite3
from datetime import UTC, datetime

# Order matters for a future import: parents before children.
TABLES = (
    "calendars",
    "sections",
    "people",
    "events",
    "event_people",
    "event_exceptions",
    "event_day_notes",
    "displays",
    "views",
)


def export_all(conn: sqlite3.Connection, version: str, now: datetime | None = None) -> dict:
    now = now or datetime.now(UTC)
    tables = {}
    for name in TABLES:
        cur = conn.execute(f"SELECT * FROM {name} ORDER BY 1")  # noqa: S608  fixed names above
        tables[name] = [dict(r) for r in cur.fetchall()]
    return {
        "format": "trmnl-calendar-export",
        "format_version": 1,
        "app_version": version,
        "exported_at": now.isoformat(timespec="seconds"),
        "tables": tables,
    }
