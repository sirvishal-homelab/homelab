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


WHO_COOKIE = "who"


def who_from_cookie(request: Request, conn: sqlite3.Connection):
    """The person this device chose, if still active. Not authentication (docs/15, M-1)."""
    from trmnl_calendar.storage.repositories import people as people_repo

    raw = request.cookies.get(WHO_COOKIE, "")
    if not raw.isdigit() or int(raw) == 0:
        return None
    try:
        person = people_repo.get_person(conn, int(raw))
    except Exception:
        return None
    return person if person.active else None


def who_chosen(request: Request) -> bool:
    """True once this device answered the question (even with 'nobody in particular')."""
    return WHO_COOKIE in request.cookies


def safe_next(url: str | None) -> str:
    """Only same-site paths, so the redirect after choosing can't leave the app."""
    if url and url.startswith("/") and not url.startswith("//") and "\\" not in url:
        return url
    return "/"


THEME_COOKIE = "theme"
THEMES = ("auto", "light", "dark")


def theme_from_cookie(request: Request) -> str:
    """This device's look: 'auto' (follow the phone), 'light' or 'dark'. Anything else is 'auto'."""
    raw = request.cookies.get(THEME_COOKIE, "auto")
    return raw if raw in THEMES else "auto"


def chrome(request: Request) -> dict:
    """Values every page needs: who this device is, the current path and the chosen look."""
    from contextlib import closing

    with closing(db.connect(request.app.state.settings.db_path)) as conn:
        return {
            "me": who_from_cookie(request, conn),
            "request_path": request.url.path,
            "theme": theme_from_cookie(request),
        }
