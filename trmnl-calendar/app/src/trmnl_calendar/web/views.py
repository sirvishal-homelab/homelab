"""Page models and template rendering. Framework-free so it can be tested without a server."""

from __future__ import annotations

import sqlite3
from datetime import date, datetime, timedelta
from functools import lru_cache

from jinja2 import Environment, PackageLoader, select_autoescape

from trmnl_calendar.domain.models import Event, Section
from trmnl_calendar.domain.recurrence import occurrences
from trmnl_calendar.domain.repeat_text import repeat_text
from trmnl_calendar.render.layouts.base import DAY_ABBR, MONTH_ABBR, fmt_range, fmt_time
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import people as people_repo
from trmnl_calendar.web import forms
from trmnl_calendar.web.render_service import today_for

UPCOMING_DAYS = 28


@lru_cache(maxsize=1)
def env() -> Environment:
    e = Environment(
        loader=PackageLoader("trmnl_calendar.web", "templates"),
        autoescape=select_autoescape(["html"]),
        trim_blocks=True,
        lstrip_blocks=True,
    )
    e.globals.update(
        repeat_choices=forms.REPEAT_CHOICES,
        freq_choices=forms.FREQ_CHOICES,
        nth_choices=forms.NTH_CHOICES,
        weekdays=forms.WEEKDAYS,
    )
    return e


def render_page(name: str, **ctx) -> str:
    return env().get_template(name).render(**ctx)


def day_label(d: date, today: date) -> str:
    if d == today:
        return "Today"
    if d == today + timedelta(days=1):
        return "Tomorrow"
    return f"{DAY_ABBR[d.weekday()]} {MONTH_ABBR[d.month - 1]} {d.day}"


def upcoming_groups(
    conn: sqlite3.Connection, now_utc: datetime, person_id: int | None = None
) -> dict:
    """The next four weeks, grouped by day, for the home page."""
    calendar = cal.get_calendar(conn)
    today = today_for(calendar.timezone, now_utc)
    end = today + timedelta(days=UPCOMING_DAYS - 1)
    events = repo.list_events(conn, today, end)
    occ = occurrences(
        events, repo.skips_for(conn, [e.id for e in events]), today, end, calendar.week_start
    )
    names = {s.id: s.name for s in cal.list_sections(conn)}
    everyone = people_repo.list_people(conn)
    person_names = {p.id: p.name for p in everyone}
    repeating = {e.id: e.freq != "none" for e in events}
    groups: dict[date, list[dict]] = {}
    for o in occ:
        if person_id is not None and o.people and person_id not in o.people:
            continue  # someone else's event; events for everyone (no people) always show
        when = ""
        if o.start_time:
            when = fmt_time(o.start_time, calendar.time_format)
            if o.end_time:
                when += "–" + fmt_time(o.end_time, calendar.time_format)
        elif o.end_time:
            when = "until " + fmt_time(o.end_time, calendar.time_format)
        trip = ""
        if o.is_multi_day:
            first = o.date - timedelta(days=o.day_number - 1)
            last = first + timedelta(days=o.day_count - 1)
            trip = f"Day {o.day_number} of {o.day_count} ({fmt_range(first, last)})"
        groups.setdefault(o.date, []).append(
            {
                "event_id": o.event_id,
                "title": o.title,
                "when": when,
                "trip": trip,
                "category": names.get(o.section_id, ""),
                "repeats": repeating.get(o.event_id, False),
                "location": o.location,
                "people": [person_names[i] for i in o.people if i in person_names],
            }
        )
    return {
        "today": today,
        "days": [(d, day_label(d, today), items) for d, items in sorted(groups.items())],
        "empty": not groups,
        "window_days": UPCOMING_DAYS,
        "people_list": [p for p in everyone if p.active],
        "filter_person": person_id,
    }


def form_context(
    conn: sqlite3.Connection,
    values: dict,
    errors: dict | None = None,
    *,
    event: Event | None = None,
    skips: list[date] | None = None,
    error_message: str = "",
) -> dict:
    sections: list[Section] = cal.list_sections(conn, include_hidden=False)
    attached = {int(i) for i in values.get("people", []) if str(i).isdigit()}
    # active people, plus any retired person already on this event so editing never drops them
    choices = [p for p in people_repo.list_people(conn) if p.active or p.id in attached]
    return {
        "people": choices,
        "values": values,
        "errors": errors or {},
        "error_message": error_message,
        "sections": sections,
        "event": event,
        "rev": event.rev if event else None,
        "repeat_text": repeat_text(event) if event else "",
        "skips": skips or [],
        "more_open": bool(
            values.get("end_date")
            or values.get("location")
            or values.get("repeat") != "none"
            or any(
                k in (errors or {})
                for k in (
                    "end_date",
                    "location",
                    "interval",
                    "by_weekday",
                    "month_nth",
                    "month_mode",
                    "until",
                    "freq",
                    "owner_id",
                )
            )
            or bool(values.get("owner_id"))
        ),
    }


def manage_categories_context(
    conn: sqlite3.Connection, values=None, errors=None, error_message=""
) -> dict:
    from trmnl_calendar.storage.repositories import sections as sec

    rows = cal.list_sections(conn)
    return {
        "categories": [(s, sec.count_events(conn, s.id)) for s in rows],
        "count": len(rows),
        "max": sec.MAX_SECTIONS,
        "values": values or {"name": "", "placement": "lane"},
        "errors": errors or {},
        "error_message": error_message,
    }


def category_edit_context(
    conn: sqlite3.Connection, section_id: int, values=None, errors=None, error_message=""
) -> dict:
    from trmnl_calendar.storage.repositories import sections as sec

    s = sec.get_section(conn, section_id)
    return {
        "category": s,
        "event_count": sec.count_events(conn, s.id),
        "others": [o for o in cal.list_sections(conn) if o.id != s.id],
        "only_one": len(cal.list_sections(conn)) <= 1,
        "values": values
        or {
            "name": s.name,
            "weight": str(s.weight),
            "visible": s.visible,
            "placement": s.placement,
        },
        "errors": errors or {},
        "error_message": error_message,
    }


def manage_people_context(
    conn: sqlite3.Connection, values=None, errors=None, error_message=""
) -> dict:
    people = people_repo.list_people(conn)
    return {
        "people": people,
        "count": len(people),
        "max": people_repo.MAX_PEOPLE,
        "values": values or {"name": "", "label": ""},
        "errors": errors or {},
        "error_message": error_message,
    }


def person_edit_context(
    conn: sqlite3.Connection, person_id: int, values=None, errors=None, error_message=""
) -> dict:
    p = people_repo.get_person(conn, person_id)
    return {
        "person": p,
        "values": values or {"name": p.name, "label": p.label, "active": p.active},
        "errors": errors or {},
        "error_message": error_message,
    }


def settings_context(
    conn: sqlite3.Connection, values=None, errors=None, error_message="", saved=False
) -> dict:
    c = cal.get_calendar(conn)
    return {
        "values": values
        or {
            "name": c.name,
            "timezone": c.timezone,
            "week_start": c.week_start,
            "time_format": c.time_format,
        },
        "errors": errors or {},
        "error_message": error_message,
        "saved": saved,
    }


def who_context(conn: sqlite3.Connection, current: int | None, next_url: str = "/") -> dict:
    return {
        "people": [p for p in people_repo.list_people(conn) if p.active],
        "current": current,
        "next_url": next_url,
    }
