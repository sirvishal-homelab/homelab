"""Page models for the Shifts pages: an overview of rotating schedules and a work month.

Framework-free like views.py. A "schedule" is a non-deleted event with a rotation (docs/20).
Each day of a schedule is one of: none (not a working day), working, in, out, off.
"""

from __future__ import annotations

import sqlite3
from datetime import date, datetime, timedelta

from trmnl_calendar.domain.recurrence import occurrences, week_floor
from trmnl_calendar.render.layouts.base import DAY_ABBR, MONTH_ABBR
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import people as people_repo
from trmnl_calendar.web.render_service import today_for
from trmnl_calendar.web.views import MONTH_NAMES

STATE_WORDS = {
    "none": "Not a working day",
    "working": "Working, not decided",
    "in": "Working, in",
    "out": "Working, out",
    "off": "Day off",
}
SHORT_WORDS = {"none": "", "working": "", "in": "in", "out": "out", "off": "off"}


def short_date(d: date) -> str:
    return f"{DAY_ABBR[d.weekday()]} {MONTH_ABBR[d.month - 1]} {d.day}"


def day_link(event_id: int, year: int, month: int, day: date) -> str:
    return f"/shifts/month?e={event_id}&y={year}&m={month}&day={day.isoformat()}"


def week_letter(event, monday: date) -> str:
    """Which week of the pattern a Monday falls in, named A, B, C, D."""
    base = week_floor(event.start_date, "mon")
    return chr(ord("A") + ((monday - base).days // 7) % len(event.rotation))


def day_states(
    conn: sqlite3.Connection, events, window_start: date, window_end: date, week_start: str
) -> dict[int, dict[date, str]]:
    """{event id: {date: state}} for every day the pattern works; other days are absent (none)."""
    ids = [e.id for e in events]
    skips = repo.skips_for(conn, ids)
    notes = repo.notes_for(conn, ids)
    out: dict[int, dict[date, str]] = {e.id: {} for e in events}
    # skips=None keeps the days off in the list, so they can be shown as days off
    for o in occurrences(events, None, window_start, window_end, week_start, notes):
        off = o.date in skips.get(o.event_id, set())
        out[o.event_id][o.date] = "off" if off else (o.note or "working")
    return out


def _schedules(conn: sqlite3.Connection, window_start: date, window_end: date):
    return [e for e in repo.list_events(conn, window_start, window_end) if e.rotation]


def _label(event, person_names: dict[int, str]) -> str:
    who = ", ".join(person_names[i] for i in event.people if i in person_names)
    return f"{who}: {event.title}" if who else event.title


def _cell(d: date, state: str, today: date, **extra) -> dict:
    label = short_date(d)
    return {
        "date": d,
        "iso": d.isoformat(),
        "day": d.day,
        "dow": DAY_ABBR[d.weekday()][0],
        "state": state,
        "word": SHORT_WORDS[state],
        "today": d == today,
        "label": f"{label}, {STATE_WORDS[state].lower()}",
        **extra,
    }


def overview_context(conn: sqlite3.Connection, now_utc: datetime, flash: str = "") -> dict:
    """One card per rotating schedule, with a strip as many weeks long as the rotation."""
    calendar = cal.get_calendar(conn)
    today = today_for(calendar.timezone, now_utc)
    monday = week_floor(today, "mon")
    probe_end = monday + timedelta(days=7 * 4 - 1)
    events = _schedules(conn, monday, probe_end)
    names = {p.id: p.name for p in people_repo.list_people(conn)}
    states = day_states(conn, events, monday, probe_end, calendar.week_start)
    cards = []
    for e in events:
        weeks = []
        for w in range(len(e.rotation)):
            first = monday + timedelta(days=7 * w)
            days = []
            for i in range(7):
                d = first + timedelta(days=i)
                days.append(
                    _cell(
                        d,
                        states[e.id].get(d, "none"),
                        today,
                        href=day_link(e.id, d.year, d.month, d) + "#day",
                    )
                )
            weeks.append(
                {
                    "name": week_letter(e, first),
                    "title": f"from {short_date(first)}",
                    "days": days,
                }
            )
        cards.append(
            {
                "event_id": e.id,
                "label": _label(e, names),
                "this_week": week_letter(e, monday),
                "weeks": weeks,
            }
        )
    return {"cards": cards, "flash": flash}


def month_context(
    conn: sqlite3.Connection,
    now_utc: datetime,
    event_id: int | None = None,
    year: int | None = None,
    month: int | None = None,
    day: str | None = None,
    me_id: int | None = None,
    flash: str = "",
) -> dict:
    """A month of one schedule: each day's state in words, and the chosen day's buttons."""
    calendar = cal.get_calendar(conn)
    today = today_for(calendar.timezone, now_utc)
    if not (year and month and 1970 <= year <= 2100 and 1 <= month <= 12):
        year, month = today.year, today.month
    first_of_month = date(year, month, 1)
    grid_start = week_floor(first_of_month, calendar.week_start)
    next_month = date(year + (month == 12), month % 12 + 1, 1)
    weeks = -(-((next_month - grid_start).days) // 7)
    grid_end = grid_start + timedelta(days=weeks * 7 - 1)

    events = _schedules(conn, grid_start, grid_end)
    names = {p.id: p.name for p in people_repo.list_people(conn)}
    chosen = next((e for e in events if e.id == event_id), None)
    if chosen is None and me_id is not None:
        chosen = next((e for e in events if me_id in e.people), None)
    if chosen is None and events:
        chosen = events[0]
    schedules = [{"event_id": e.id, "label": _label(e, names), "on": e is chosen} for e in events]
    title = f"{MONTH_NAMES[month - 1]} {year}"
    prev_month = first_of_month - timedelta(days=1)
    base = {
        "year": year,
        "month": month,
        "title": title,
        "schedules": schedules,
        "chosen": chosen,
        "prev": (prev_month.year, prev_month.month),
        "next": (next_month.year, next_month.month),
        "is_current": (year, month) == (today.year, today.month),
        "flash": flash,
        "rows": [],
        "weekday_names": [],
        "selected": None,
    }
    if chosen is None:
        return base

    states = day_states(conn, [chosen], grid_start, grid_end, calendar.week_start)[chosen.id]
    selected = None
    if day:
        try:
            selected = date.fromisoformat(day)
        except ValueError:
            selected = None
    rows = []
    for w in range(weeks):
        cells = []
        for i in range(7):
            d = grid_start + timedelta(days=w * 7 + i)
            c = _cell(d, states.get(d, "none"), today, in_month=d.month == month)
            c["selected"] = d == selected
            c["href"] = day_link(chosen.id, year, month, d) + "#day"
            cells.append(c)
        rows.append(cells)
    panel = None
    if selected:
        state = states.get(selected, "none")
        panel = {
            "event_id": chosen.id,
            "title": chosen.title,
            "date": selected.isoformat(),
            "scheduled": state not in ("none", "off"),
            "note": state if state in ("in", "out") else "",
            "state": state,
            "words": STATE_WORDS[state],
            "heading": short_date(selected),
        }
    base.update(
        {
            "rows": rows,
            "weekday_names": [
                DAY_ABBR[(grid_start + timedelta(days=i)).weekday()] for i in range(7)
            ],
            "selected": selected,
            "panel": panel,
            "back": day_link(chosen.id, year, month, selected or first_of_month),
        }
    )
    return base
