"""Build the render context from the database and render the active view (slice 1: Week - lanes)."""

from __future__ import annotations

import sqlite3
from datetime import UTC, date, datetime
from zoneinfo import ZoneInfo

from trmnl_calendar.domain.recurrence import occurrences
from trmnl_calendar.render.layouts import week_lanes
from trmnl_calendar.render.layouts.base import PlanContext
from trmnl_calendar.render.pipeline import RenderCache, RenderResult, render
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo

PROFILE = {"width": 800, "height": 480, "bit_depth": 2}  # the display profile arrives in slice 4


def today_for(timezone: str, now_utc: datetime) -> date:
    return now_utc.astimezone(ZoneInfo(timezone)).date()


def render_current(
    conn: sqlite3.Connection, now_utc: datetime, cache: RenderCache | None = None
) -> RenderResult:
    calendar = cal.get_calendar(conn)
    today = today_for(calendar.timezone, now_utc)
    first, last = week_lanes.window(today, calendar.week_start)
    events = repo.list_events(conn, first, last)
    occ = occurrences(
        events, repo.skips_for(conn, [e.id for e in events]), first, last, calendar.week_start
    )
    sections = tuple(s for s in cal.list_sections(conn, include_hidden=False))
    visible_ids = {s.id for s in sections}
    ctx = PlanContext(
        today=today,
        sections=sections,
        occurrences=tuple(o for o in occ if o.section_id in visible_ids),
        week_start=calendar.week_start,
        time_format=calendar.time_format,
        **PROFILE,
    )
    result = render(ctx)
    if cache is not None:
        cache.put(result.hash, result.png)
    return result


def now_utc() -> datetime:
    return datetime.now(UTC)
