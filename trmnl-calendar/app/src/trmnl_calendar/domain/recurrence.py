"""Occurrence expansion (docs/03 section 4, ADR-0010). Pure functions."""

from __future__ import annotations

import calendar
from collections.abc import Iterable, Iterator
from datetime import date, timedelta

from trmnl_calendar.domain.models import WEEKDAY_CODES, Event, Occurrence

MAX_WINDOW_DAYS = 400


def week_floor(d: date, week_start: str = "sun") -> date:
    """First day of the week containing d."""
    back = (d.weekday() + 1) % 7 if week_start == "sun" else d.weekday()
    return d - timedelta(days=back)


def _clamp_day(year: int, month: int, day: int) -> date:
    return date(year, month, min(day, calendar.monthrange(year, month)[1]))


def nth_weekday(year: int, month: int, weekday: int, nth: int) -> date | None:
    """nth (1-4) or last (-1) given weekday (0=Monday) of the month."""
    last = calendar.monthrange(year, month)[1]
    days = [
        date(year, month, d)
        for d in range(1, last + 1)
        if date(year, month, d).weekday() == weekday
    ]
    if nth == -1:
        return days[-1]
    return days[nth - 1] if 1 <= nth <= len(days) else None


def _add_months(year: int, month: int, n: int) -> tuple[int, int]:
    idx = year * 12 + (month - 1) + n
    return idx // 12, idx % 12 + 1


def _starts(ev: Event, lo: date, hi: date, week_start: str) -> Iterator[date]:
    """Start dates of the series' occurrences with lo <= start <= hi (before until/skips)."""
    s = ev.start_date
    if hi < s:
        return
    lo = max(lo, s)
    if ev.freq == "none":
        if lo <= s <= hi:
            yield s
    elif ev.freq == "daily":
        k = max(0, -(-(lo - s).days // ev.interval))
        d = s + timedelta(days=k * ev.interval)
        while d <= hi:
            yield d
            d += timedelta(days=ev.interval)
    elif ev.freq == "weekly":
        codes = ev.by_weekday or (WEEKDAY_CODES[s.weekday()],)
        wanted = {WEEKDAY_CODES.index(c) for c in codes}
        base = week_floor(s, week_start)
        d = lo
        while d <= hi:
            weeks = (week_floor(d, week_start) - base).days // 7
            if d.weekday() in wanted and weeks % ev.interval == 0:
                yield d
            d += timedelta(days=1)
    elif ev.freq == "monthly":
        k = 0
        while True:
            y, m = _add_months(s.year, s.month, k * ev.interval)
            if date(y, m, 1) > hi:
                return
            if ev.month_mode == "weekday":
                wd = WEEKDAY_CODES.index(ev.by_weekday[0]) if ev.by_weekday else s.weekday()
                d = nth_weekday(y, m, wd, ev.month_nth or 1)
            else:
                d = _clamp_day(y, m, s.day)
            if d is not None and lo <= d <= hi and d >= s:
                yield d
            k += 1
    elif ev.freq == "yearly":
        k = 0
        while True:
            y = s.year + k * ev.interval
            if date(y, 1, 1) > hi:
                return
            d = _clamp_day(y, s.month, s.day)
            if lo <= d <= hi:
                yield d
            k += 1


def occurrences(
    events: Iterable[Event],
    skips: dict[int, set[date]] | None,
    window_start: date,
    window_end: date,
    week_start: str = "sun",
    people: dict[int, tuple[int, ...]] | None = None,
) -> list[Occurrence]:
    """Per-day occurrences inside [window_start, window_end], in display order."""
    if window_end < window_start:
        return []
    if (window_end - window_start).days > MAX_WINDOW_DAYS:
        raise ValueError(f"window longer than {MAX_WINDOW_DAYS} days")
    skips = skips or {}
    people = people or {}
    out: list[Occurrence] = []
    for ev in events:
        dur = ev.duration_days
        for start in _starts(ev, window_start - timedelta(days=dur), window_end, week_start):
            if ev.until is not None and start > ev.until:
                break
            if start in skips.get(ev.id, ()):
                continue
            for i in range(dur + 1):
                day = start + timedelta(days=i)
                if not window_start <= day <= window_end:
                    continue
                first, last = i == 0, i == dur
                out.append(
                    Occurrence(
                        event_id=ev.id,
                        title=ev.title,
                        location=ev.location,
                        section_id=ev.section_id,
                        date=day,
                        start_time=ev.start_time if first else None,
                        end_time=ev.end_time if last else None,
                        all_day=ev.start_time is None,
                        day_number=i + 1,
                        day_count=dur + 1,
                        series_start=start,
                        people=people.get(ev.id, ()),
                    )
                )
    out.sort(key=sort_key)
    return out


def sort_key(o: Occurrence) -> tuple:
    """Date, all-day first, start time, title (case-insensitive), event id."""
    return (o.date, 0 if o.all_day else 1, o.start_time or "", o.title.lower(), o.event_id)
