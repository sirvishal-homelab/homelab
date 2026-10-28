"""Build the render context from the database and render a view (docs/04 section 2).

The preview and the push both come through here, so they cannot drift apart.
"""

from __future__ import annotations

import sqlite3
from datetime import UTC, date, datetime
from zoneinfo import ZoneInfo

from trmnl_calendar.domain.models import View
from trmnl_calendar.domain.recurrence import occurrences
from trmnl_calendar.render.layouts.base import PlanContext
from trmnl_calendar.render.layouts.registry import LAYOUTS
from trmnl_calendar.render.pipeline import RenderCache, RenderResult, render
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import people as people_repo
from trmnl_calendar.storage.repositories import views as views_repo


def today_for(timezone: str, now_utc: datetime) -> date:
    return now_utc.astimezone(ZoneInfo(timezone)).date()


def render_view(
    conn: sqlite3.Connection,
    view: View,
    now_utc: datetime,
    cache: RenderCache | None = None,
    as_of: date | None = None,
) -> RenderResult:
    calendar = cal.get_calendar(conn)
    display = views_repo.get_display(conn)
    today = as_of or today_for(calendar.timezone, now_utc)
    layout = LAYOUTS[view.layout]
    first, last = layout.window(today, calendar.week_start)
    events = repo.list_events(conn, first, last)
    occ = occurrences(
        events, repo.skips_for(conn, [e.id for e in events]), first, last, calendar.week_start
    )
    sections = tuple(s for s in cal.list_sections(conn, include_hidden=False))
    if view.categories_mode == "only":
        sections = tuple(s for s in sections if s.id in view.category_ids)
    shown = {s.id for s in sections}
    wanted = set(view.person_ids) if view.people_mode == "only" else None
    chosen = tuple(
        o
        for o in occ
        if o.section_id in shown and (wanted is None or not o.people or wanted & set(o.people))
    )
    ctx = PlanContext(
        today=today,
        sections=sections,
        occurrences=chosen,
        week_start=calendar.week_start,
        time_format=calendar.time_format,
        weekdays=view.weekdays,
        options=view.options,
        people={p.id: p for p in people_repo.list_people(conn)},
        width=display.width,
        height=display.height,
        bit_depth=display.bit_depth,
    )
    result = render(ctx, view.layout)
    if cache is not None:
        cache.put(result.hash, result.png)
    return result


def render_current(
    conn: sqlite3.Connection, now_utc: datetime, cache: RenderCache | None = None
) -> RenderResult:
    """The image for the display's active view."""
    return render_view(conn, views_repo.active_view(conn), now_utc, cache)


def now_utc() -> datetime:
    return datetime.now(UTC)
