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
                    "count",
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


def local_when(iso: str | None, timezone: str, time_format: str) -> str:
    """'Oct 28, 10:44a' from a stored UTC timestamp, in the calendar's timezone."""
    if not iso:
        return ""
    from zoneinfo import ZoneInfo

    t = datetime.strptime(iso, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=ZoneInfo("UTC"))
    t = t.astimezone(ZoneInfo(timezone))
    return f"{MONTH_ABBR[t.month - 1]} {t.day}, {fmt_time(t.strftime('%H:%M'), time_format)}"


def screen_context(conn: sqlite3.Connection, settings, scheduler) -> dict:
    from trmnl_calendar.publish import check, status

    c = cal.get_calendar(conn)
    st = status.screen_status(conn, settings, scheduler)
    when = lambda k: local_when(st[k], c.timezone, c.time_format)  # noqa: E731
    return {
        **st,
        "missing": check.missing_settings(settings),
        "pushed_when": when("last_pushed_at"),
        "attempt_when": when("last_attempt_at"),
        "next_when": when("next_check_at"),
        "interval_minutes": settings.push_interval_seconds // 60,
        "device_minutes": -(-st["device_refresh_seconds"] // 60)
        if st["device_refresh_seconds"]
        else None,
    }


def settings_context(
    conn: sqlite3.Connection,
    values=None,
    errors=None,
    error_message="",
    saved=False,
    screen=None,
    check_steps=None,
    screen_flash="",
) -> dict:
    c = cal.get_calendar(conn)
    return {
        "screen": screen,
        "check_steps": check_steps,
        "screen_flash": screen_flash,
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


def views_list_context(conn: sqlite3.Connection, error_message: str = "") -> dict:
    from trmnl_calendar.domain.view_templates import TEMPLATES
    from trmnl_calendar.render.layouts import registry
    from trmnl_calendar.storage.repositories import views as views_repo

    active = views_repo.active_view(conn)
    return {
        "views": [(v, registry.LAYOUTS[v.layout].LABEL) for v in views_repo.list_views(conn)],
        "active_id": active.id,
        "templates": [(tid, t["label"], t["help"]) for tid, t in TEMPLATES.items()],
        "error_message": error_message,
    }


def view_editor_context(
    conn: sqlite3.Connection,
    now_utc: datetime,
    cache,
    values: dict,
    view_id: int | None = None,
    errors: dict | None = None,
    error_message: str = "",
    flash: str = "",
) -> dict:
    """Everything the editor page needs, including a server-rendered preview (works without JS)."""
    from trmnl_calendar.domain.models import WEEKDAY_CODES, WEEKDAY_NAMES
    from trmnl_calendar.errors import ValidationFailed
    from trmnl_calendar.render.layouts import registry
    from trmnl_calendar.storage.repositories import views as views_repo
    from trmnl_calendar.web import render_service

    image = None
    fit: list = []
    try:
        preview_view = views_repo.build_view(conn, {**values, "name": "Preview"})
        result = render_service.render_view(conn, preview_view, now_utc, cache)
        image = {
            "url": f"/render/{result.hash}.png",
            "width": result.width,
            "height": result.height,
        }
        fit = [f for f in result.fit.findings if f.code != "titles_truncated"]
        cut = [f for f in result.fit.findings if f.code == "titles_truncated"]
        fit = fit + cut
    except ValidationFailed:
        pass
    display = views_repo.get_display(conn)
    layouts = [
        {
            "id": lid,
            "label": m.LABEL,
            "options": list(m.OPTIONS),
        }
        for lid, m in registry.LAYOUTS.items()
    ]
    return {
        "view_id": view_id,
        "is_active": bool(view_id) and views_repo.active_view(conn).id == view_id,
        "values": values,
        "errors": errors or {},
        "error_message": error_message,
        "flash": flash,
        "image": image,
        "fit": fit,
        "profile": display,
        "layouts": layouts,
        "option_labels": registry.OPTION_LABELS,
        "option_order": list(registry.OPTION_DEFAULTS),
        "option_defaults": registry.OPTION_DEFAULTS,
        "weekdays": list(zip(WEEKDAY_CODES, WEEKDAY_NAMES, strict=True)),
        "categories": cal.list_sections(conn),
        "people_all": [p for p in people_repo.list_people(conn)],
    }


def holidays_context(
    conn: sqlite3.Connection,
    region: str,
    years: list[int],
    this_year: int,
    category_id: int | None = None,
    done: str = "",
    error_message: str = "",
) -> dict:
    """The holiday picker: regions, years, and (once chosen) the list with what is already added."""
    from trmnl_calendar.domain import holidays as H
    from trmnl_calendar.storage.repositories import calendar as cal_repo

    sections = cal_repo.list_sections(conn)
    target = next((s for s in sections if s.id == category_id), None) or next(
        (s for s in sections if s.name.lower() == "holidays"), None
    )
    groups = []
    if region in H.REGIONS and years:
        items = [h for y in years for h in H.holidays_for(region, y)]
        for gid, label in H.GROUPS:
            rows = []
            for h in items:
                if h.group != gid:
                    continue
                added = bool(
                    target
                    and conn.execute(
                        "SELECT 1 FROM events WHERE calendar_id = ? AND section_id = ? "
                        "AND deleted_at IS NULL AND start_date = ? AND title = ? COLLATE NOCASE",
                        (cal_repo.CALENDAR_ID, target.id, h.date.isoformat(), h.name),
                    ).fetchone()
                )
                rows.append(
                    {
                        "token": f"{h.key}:{h.date.year}",
                        "name": h.name,
                        "when": (
                            f"{DAY_ABBR[h.date.weekday()]} {MONTH_ABBR[h.date.month - 1]} "
                            f"{h.date.day}, {h.date.year}"
                        ),
                        "checked": gid in H.CHECKED_BY_DEFAULT and not added,
                        "added": added,
                    }
                )
            if rows:
                groups.append({"id": gid, "label": label, "rows": rows})
    return {
        "regions": H.region_choices(),
        "region": region if region in H.REGIONS else "",
        "region_note": H.REGIONS[region]["note"] if region in H.REGIONS else "",
        "year_choices": [this_year, this_year + 1, this_year + 2],
        "years": years,
        "groups": groups,
        "sections": sections,
        "category_id": category_id,
        "target_name": target.name if target else "Holidays",
        "target_exists": target is not None,
        "done": done,
        "error_message": error_message,
    }
