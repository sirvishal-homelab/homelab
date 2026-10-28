"""HTML form <-> event data. Pure: no web framework imports, so it is unit-tested directly."""

from __future__ import annotations

from urllib.parse import parse_qs

from trmnl_calendar.domain.models import WEEKDAY_CODES, WEEKDAY_NAMES, Event
from trmnl_calendar.render.layouts.base import OPTION_DEFAULTS

REPEAT_CHOICES = (
    ("none", "Does not repeat"),
    ("daily", "Every day"),
    ("weekly", "Every week"),
    ("biweekly", "Every 2 weeks"),
    ("monthly", "Every month"),
    ("yearly", "Every year"),
    ("custom", "Custom..."),
)
FREQ_CHOICES = (("daily", "days"), ("weekly", "weeks"), ("monthly", "months"), ("yearly", "years"))
NTH_CHOICES = (("1", "first"), ("2", "second"), ("3", "third"), ("4", "fourth"), ("-1", "last"))
WEEKDAYS = tuple(zip(WEEKDAY_CODES, WEEKDAY_NAMES, strict=True))


def parse_body(body: bytes) -> dict[str, list[str]]:
    return parse_qs(body.decode("utf-8", "replace"), keep_blank_values=True)


def _one(f: dict[str, list[str]], key: str) -> str:
    v = f.get(key)
    return v[0].strip() if v else ""


def form_to_event_data(f: dict[str, list[str]]) -> dict:
    """Turn submitted form fields into the dict validate_event expects."""
    sid = _one(f, "section_id")
    data: dict = {
        "title": _one(f, "title"),
        "location": _one(f, "location"),
        "section_id": int(sid) if sid.isdigit() else None,
        "start_date": _one(f, "start_date"),
        "end_date": _one(f, "end_date") or None,
        "start_time": _one(f, "start_time") or None,
        "end_time": _one(f, "end_time") or None,
        "freq": "none",
        "interval": 1,
        "by_weekday": [],
        "month_mode": None,
        "month_nth": None,
        "until": None,
        "count": None,
        "people_ids": [int(v) for v in f.get("people", []) if v.isdigit()],
        "owner_id": int(_one(f, "owner_id")) if _one(f, "owner_id").isdigit() else None,
    }
    choice = _one(f, "repeat") or "none"
    if choice in ("daily", "weekly", "monthly", "yearly"):
        data["freq"] = choice
    elif choice == "biweekly":
        data["freq"], data["interval"] = "weekly", 2
    elif choice == "custom":
        data["freq"] = _one(f, "custom_freq") or "weekly"
        data["interval"] = _one(f, "custom_interval") or 1
        data["by_weekday"] = [c for c in f.get("by_weekday", []) if c]
        data["month_mode"] = _one(f, "month_mode") or None
        data["month_nth"] = _one(f, "month_nth") or None
        ends = _one(f, "ends") or "never"
        data["until"] = (_one(f, "until") or None) if ends == "until" else None
        data["count"] = (_one(f, "count") or None) if ends == "count" else None
    return data


def event_to_values(
    ev: Event | None,
    default_date: str = "",
    default_section: int | None = None,
    default_people: tuple[int, ...] = (),
) -> dict:
    """Prefill values (all strings) for the form; the repeat choice is inferred from the rule."""
    if ev is None:
        return {
            "title": "",
            "location": "",
            "section_id": str(default_section or ""),
            "start_date": default_date,
            "end_date": "",
            "start_time": "",
            "end_time": "",
            "repeat": "none",
            "custom_freq": "weekly",
            "custom_interval": "1",
            "by_weekday": [],
            "month_mode": "date",
            "month_nth": "1",
            "until": "",
            "ends": "never",
            "count": "",
            "people": [str(i) for i in default_people],
            "owner_id": "",
        }
    own_day = WEEKDAY_CODES[ev.start_date.weekday()]
    plain_weekly = ev.freq == "weekly" and ev.by_weekday in ((), (own_day,))
    simple = ev.until is None and ev.count is None and ev.month_mode in (None, "date")
    if ev.freq == "none":
        choice = "none"
    elif simple and ev.interval == 1 and ev.freq in ("daily", "monthly", "yearly"):
        choice = ev.freq
    elif simple and plain_weekly and ev.interval == 1:
        choice = "weekly"
    elif simple and plain_weekly and ev.interval == 2:
        choice = "biweekly"
    else:
        choice = "custom"
    return {
        "title": ev.title,
        "location": ev.location,
        "section_id": str(ev.section_id),
        "start_date": ev.start_date.isoformat(),
        "end_date": ev.end_date.isoformat() if ev.end_date else "",
        "start_time": ev.start_time or "",
        "end_time": ev.end_time or "",
        "repeat": choice,
        "custom_freq": ev.freq if ev.freq != "none" else "weekly",
        "custom_interval": str(ev.interval),
        "by_weekday": list(ev.by_weekday),
        "month_mode": ev.month_mode or "date",
        "month_nth": str(ev.month_nth or 1),
        "until": ev.until.isoformat() if ev.until else "",
        "ends": "until" if ev.until else ("count" if ev.count else "never"),
        "count": str(ev.count) if ev.count else "",
        "people": [str(i) for i in ev.people],
        "owner_id": str(ev.owner_id) if ev.owner_id else "",
    }


def values_from_form(f: dict[str, list[str]]) -> dict:
    """Echo submitted values back into the form after a validation error."""
    v = event_to_values(None)
    for k in v:
        if k in ("by_weekday", "people"):
            v[k] = [c for c in f.get(k, []) if c]
        elif k in f:
            v[k] = _one(f, k)
    return v


def view_form_to_dict(f: dict[str, list[str]]) -> dict:
    """Turn the view editor's fields into the dict build_view validates."""

    def ids(key: str) -> list[int]:
        return [int(v) for v in f.get(key, []) if v.isdigit()]

    return {
        "name": _one(f, "name"),
        "template": _one(f, "template") or "blank",
        "layout": _one(f, "layout"),
        "weekdays": [c for c in f.get("weekday", []) if c],
        "categories": {"mode": _one(f, "categories_mode") or "all", "ids": ids("category")},
        "people": {"mode": _one(f, "people_mode") or "all", "ids": ids("person")},
        "options": {k: k in f for k in OPTION_DEFAULTS},
    }
