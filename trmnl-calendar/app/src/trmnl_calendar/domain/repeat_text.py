"""Plain-language description of a repeat rule, for the form and API (docs/17 section 2)."""

from __future__ import annotations

from datetime import date

from trmnl_calendar.domain.models import WEEKDAY_CODES, WEEKDAY_NAMES, Event

_NTH = {1: "first", 2: "second", 3: "third", 4: "fourth", -1: "last"}
_UNIT = {
    "daily": ("day", "days"),
    "weekly": ("week", "weeks"),
    "monthly": ("month", "months"),
    "yearly": ("year", "years"),
}


def _long_date(d: date) -> str:
    return f"{d.day} {d.strftime('%B %Y')}"


def repeat_text(ev: Event) -> str:
    if ev.freq == "none":
        return "Does not repeat"
    one, many = _UNIT[ev.freq]
    text = f"Repeats every {one}" if ev.interval == 1 else f"Repeats every {ev.interval} {many}"
    if ev.freq == "weekly":
        codes = ev.by_weekday or (WEEKDAY_CODES[ev.start_date.weekday()],)
        names = [WEEKDAY_NAMES[WEEKDAY_CODES.index(c)] for c in codes]
        text += " on " + (
            names[0] if len(names) == 1 else ", ".join(names[:-1]) + " and " + names[-1]
        )
    elif ev.freq == "monthly" and ev.month_mode == "weekday" and ev.by_weekday:
        name = WEEKDAY_NAMES[WEEKDAY_CODES.index(ev.by_weekday[0])]
        text += f" on the {_NTH.get(ev.month_nth or 1, 'first')} {name}"
    elif ev.freq == "monthly":
        text += f" on day {ev.start_date.day}"
    if ev.until:
        text += f", until {_long_date(ev.until)}"
    return text
