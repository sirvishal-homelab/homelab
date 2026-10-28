"""Per-request dependencies."""

from __future__ import annotations

import sqlite3
from collections.abc import Iterator
from datetime import datetime

from fastapi import Request

from trmnl_calendar.errors import PreconditionRequired
from trmnl_calendar.storage import db


def get_conn(request: Request) -> Iterator[sqlite3.Connection]:
    conn = db.connect(request.app.state.settings.db_path)
    try:
        yield conn
    finally:
        conn.close()


def now_utc(request: Request) -> datetime:
    return request.app.state.clock()


def parse_if_match(value: str | None) -> int:
    """If-Match carries the event's rev. Missing or unreadable means the caller must reload."""
    if value is None:
        raise PreconditionRequired("Reload the event and try again (If-Match is required).")
    raw = value.strip()
    if raw.startswith("W/"):
        raw = raw[2:]
    try:
        return int(raw.strip('"'))
    except ValueError:
        raise PreconditionRequired("If-Match must be the event's rev number.") from None
