"""Event validation (docs/03). Returns cleaned values or field errors in plain language."""

from __future__ import annotations

import re
from datetime import date

from trmnl_calendar.domain.models import FREQS, WEEKDAY_CODES

_TIME = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")
MIN_DATE, MAX_DATE = date(1970, 1, 1), date(2100, 12, 31)


def parse_date(value: object) -> date | None:
    if isinstance(value, date):
        return value
    if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        try:
            return date.fromisoformat(value)
        except ValueError:
            return None
    return None


def _blank(v: object) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def validate_event(data: dict) -> tuple[dict, dict[str, str]]:
    """Validate a complete event. Returns (clean, errors); clean only counts if errors is empty."""
    err: dict[str, str] = {}
    clean: dict = {}

    title = (data.get("title") or "").strip() if isinstance(data.get("title"), str) else ""
    if not title:
        err["title"] = "Enter a short description"
    elif len(title) > 120:
        err["title"] = "Keep the description under 120 characters"
    clean["title"] = title

    loc = data.get("location") or ""
    loc = loc.strip() if isinstance(loc, str) else ""
    if len(loc) > 120:
        err["location"] = "Keep the location under 120 characters"
    clean["location"] = loc

    sid = data.get("section_id")
    if not isinstance(sid, int) or isinstance(sid, bool):
        err["section_id"] = "Choose a category"
    clean["section_id"] = sid

    start = parse_date(data.get("start_date"))
    if start is None:
        err["start_date"] = "Choose a date"
    elif not MIN_DATE <= start <= MAX_DATE:
        err["start_date"] = "Choose a date between 1970 and 2100"
    clean["start_date"] = start

    end = None
    if not _blank(data.get("end_date")):
        end = parse_date(data.get("end_date"))
        if end is None:
            err["end_date"] = "Choose a valid end date"
        elif not MIN_DATE <= end <= MAX_DATE:
            err["end_date"] = "Choose a date between 1970 and 2100"
        elif start is not None and end < start:
            err["end_date"] = "The end date can't be before the start date"
        elif start is not None and end == start:
            end = None  # a one-day "trip" is just an event
    clean["end_date"] = end

    times = {}
    for key, label in (("start_time", "start"), ("end_time", "end")):
        v = data.get(key)
        if _blank(v):
            times[key] = None
        elif isinstance(v, str) and _TIME.match(v.strip()):
            times[key] = v.strip()
        else:
            times[key] = None
            err[key] = f"Enter the {label} time like 09:30"
    if times["end_time"] and not times["start_time"]:
        err.setdefault("end_time", "Add a start time first, or clear the end time")
    if (
        times["start_time"]
        and times["end_time"]
        and end is None
        and times["end_time"] <= times["start_time"]
    ):
        err["end_time"] = "The end time must be after the start time"
    clean.update(times)

    freq = data.get("freq") or "none"
    interval = data.get("interval", 1)
    if interval in (None, ""):
        interval = 1
    by_weekday = tuple(data.get("by_weekday") or ())
    month_mode = data.get("month_mode") or None
    month_nth = data.get("month_nth")
    if month_nth in ("", None):
        month_nth = None
    until = None
    if freq not in FREQS:
        err["freq"] = "Choose how often it repeats"
        freq = "none"
    try:
        interval = int(interval)
    except (TypeError, ValueError):
        err["interval"] = "Enter a number from 1 to 99"
        interval = 1
    if not 1 <= interval <= 99:
        err["interval"] = "Enter a number from 1 to 99"
        interval = 1
    bad_days = [c for c in by_weekday if c not in WEEKDAY_CODES]
    if bad_days:
        err["by_weekday"] = "Choose weekdays from the list"
        by_weekday = ()
    by_weekday = tuple(sorted(set(by_weekday), key=WEEKDAY_CODES.index)) if not bad_days else ()

    if freq == "none":
        interval, by_weekday, month_mode, month_nth = 1, (), None, None
    else:
        if not _blank(data.get("until")):
            until = parse_date(data.get("until"))
            if until is None:
                err["until"] = "Choose a valid end date for the repeat"
            elif start is not None and until < start:
                err["until"] = "The repeat can't end before the event starts"
        if freq == "weekly":
            month_mode, month_nth = None, None
        elif freq == "monthly":
            month_mode = month_mode or "date"
            if month_mode not in ("date", "weekday"):
                err["month_mode"] = "Choose a date or a weekday"
                month_mode = "date"
            if month_mode == "weekday":
                try:
                    month_nth = int(month_nth) if month_nth is not None else None
                except (TypeError, ValueError):
                    month_nth = None
                if month_nth not in (1, 2, 3, 4, -1):
                    err["month_nth"] = "Choose first, second, third, fourth or last"
                if len(by_weekday) != 1:
                    err["by_weekday"] = "Choose one weekday"
            else:
                by_weekday, month_nth = (), None
        else:  # daily, yearly
            by_weekday, month_mode, month_nth = (), None, None
    clean.update(
        freq=freq,
        interval=interval,
        by_weekday=by_weekday,
        month_mode=month_mode,
        month_nth=month_nth,
        until=until,
    )
    return clean, err
