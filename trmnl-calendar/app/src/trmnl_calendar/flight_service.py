"""Flight status: look flights up on their schedule, and add the status word to occurrences.

docs/21 and ADR-0015. Lookups run in the app's flight loop (main.py) and from the event page's
"Check now" button; drawing only reads what is stored, so the image stays a function of the data.
"""

from __future__ import annotations

import logging
import sqlite3
from dataclasses import dataclass, replace
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from trmnl_calendar.domain import flights as F
from trmnl_calendar.domain.models import Event, Occurrence
from trmnl_calendar.domain.recurrence import occurrences
from trmnl_calendar.errors import EditConflict, ValidationFailed
from trmnl_calendar.flight_source import FlightSource, FlightSourceError
from trmnl_calendar.render.layouts.base import fmt_time
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as events_repo
from trmnl_calendar.storage.repositories import flights as repo

log = logging.getLogger(__name__)
CREDIT = "Flight data: AeroDataBox"


def annotate(conn: sqlite3.Connection, occ: list[Occurrence]) -> list[Occurrence]:
    """Add "(On time)", "(Delayed)" ... to the title of each occurrence holding a known flight."""
    keys = {(c, o.date) for o in occ for c in F.codes_for_day(o)}
    if not keys:
        return occ
    known = repo.infos_for(conn, keys)
    if not known:
        return occ
    show = _formatter(conn)
    out = []
    for o in occ:
        codes = F.codes_for_day(o)
        if codes:
            _, info = F.pick(codes, {c: known[(c, o.date)] for c in codes if (c, o.date) in known})
            w = F.label(info, show)
            if w:
                o = replace(o, title=f"{o.title} ({w})", flight=w)
        out.append(o)
    return out


def _formatter(conn: sqlite3.Connection):
    style = cal.get_calendar(conn).time_format
    return lambda hhmm: fmt_time(hhmm, style)


def event_flights(event: Event) -> list[tuple[str, date]]:
    """(code, day) pairs for the event's own dates, for its page (same rules as codes_for_day)."""
    codes = F.find_codes(f"{event.title} {event.location}")
    if not codes:
        return []
    if not event.end_date or event.end_date == event.start_date:
        return [(c, event.start_date) for c in codes]
    pairs = [(codes[0], event.start_date)]
    if len(codes) > 1:
        pairs.append((codes[-1], event.end_date))
    return pairs


@dataclass
class CheckResult:
    code: str
    day: date
    word: str
    info: F.FlightInfo | None
    error: str = ""
    checked_at: datetime | None = None


class FlightRefresher:
    def __init__(self, source: FlightSource, monthly_limit: int):
        self.source = source
        self.limit = monthly_limit
        self.capped = False

    def _today(self, conn: sqlite3.Connection, now: datetime) -> date:
        return now.astimezone(ZoneInfo(cal.get_calendar(conn).timezone)).date()

    def _lookup(self, conn: sqlite3.Connection, code: str, day: date, now: datetime) -> bool:
        """Ask once, store the answer, fill a missing start time. True if the screen changes."""
        show = _formatter(conn)
        old = F.label(repo.info(repo.get(conn, code, day)), show)
        if repo.calls_this_month(conn, now) >= self.limit:
            self.capped = True
            return False
        repo.count_call(conn, now)
        try:
            info = self.source.lookup(code, day)
        except FlightSourceError as exc:
            log.warning("flight lookup %s %s failed: %s", code, day, exc)
            repo.save_error(conn, code, day, str(exc), now, now + F.RETRY_AFTER_ERROR)
            return False
        repo.save(conn, info, now, F.next_due(info, now))
        filled = fill_missing_time(conn, code, day, info)
        return filled or old != F.label(info, show)  # a later new time is a change too

    def due_pairs(self, conn: sqlite3.Connection, now: datetime) -> list[tuple[str, date]]:
        """Flights on the calendar from yesterday to tomorrow that should be asked about now."""
        today = self._today(conn, now)
        first, last = today - timedelta(days=1), today + timedelta(days=1)
        evs = events_repo.list_events(conn, first, last)
        ids = [e.id for e in evs]
        occ = occurrences(evs, events_repo.skips_for(conn, ids), first, last)
        pairs: list[tuple[str, date]] = []
        for o in occ:
            for code in F.codes_for_day(o):
                key = (code, o.date)
                if key in pairs:
                    continue
                row = repo.get(conn, code, o.date)
                if row is None:
                    if F.first_check_due(o.date, today):
                        pairs.append(key)
                elif repo.is_due(row, now):
                    pairs.append(key)
        return pairs

    def refresh(self, conn: sqlite3.Connection, now: datetime) -> bool:
        """One pass: purge old answers, ask about due flights. True if a shown word changed."""
        repo.purge(conn, now)
        self.capped = repo.calls_this_month(conn, now) >= self.limit
        changed = False
        for code, day in self.due_pairs(conn, now):
            changed |= self._lookup(conn, code, day, now)
            if self.capped:
                break
        return changed

    def check_now(self, conn: sqlite3.Connection, event: Event, now: datetime) -> list[CheckResult]:
        """The event page's button: ask about this event's flights now, whatever the schedule."""
        out = []
        for code, day in event_flights(event):
            self._lookup(conn, code, day, now)
            out.append(stored(conn, code, day))
        return out


def fill_missing_time(conn: sqlite3.Connection, code: str, day: date, info: F.FlightInfo) -> bool:
    """Give a flight event with no start time the planned departure (owner, 2026-11-15; docs/21).

    Only an event's first flight on its first day, only one-off events, and never a time someone
    typed. Saved like an edit, so it shows on the event page and can be changed back.
    """
    if not info.found or not info.sched_dep_local:
        return False
    filled = False
    for ev in events_repo.list_events(conn, day, day):
        if ev.start_time is not None or ev.freq != "none" or ev.start_date != day:
            continue
        if event_flights(ev)[:1] != [(code, day)]:
            continue
        try:
            events_repo.update_event(conn, ev.id, {"start_time": info.sched_dep_local}, ev.rev)
        except (EditConflict, ValidationFailed):
            continue  # an edit in progress, or a time the event cannot take: leave it
        log.info("event %d: start time set to %s from flight %s", ev.id, info.sched_dep_local, code)
        filled = True
    return filled


def stored(conn: sqlite3.Connection, code: str, day: date) -> CheckResult:
    row = repo.get(conn, code, day)
    info = repo.info(row)
    if row is not None and row["last_error"] and not row["status"]:
        info = None  # only a failure is stored: nothing known about the flight yet
    return CheckResult(
        code=code,
        day=day,
        word=F.word(info),
        info=info,
        error=row["last_error"] if row is not None else "",
        checked_at=repo.checked_at(row) if row is not None else None,
    )
