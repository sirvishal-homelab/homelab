"""Flight codes in event text, the word shown for a flight, and when to look it up again (docs/21).

Pure: no database, no network, no clock. Callers pass `now` (UTC) in.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import UTC, date, datetime, timedelta

from trmnl_calendar.domain.models import Occurrence

# An airline code (two letters, or a letter and a digit either way round), then 1 to 4 digits.
# Upper case only, so ordinary words in a title do not match (docs/21 section 2).
_CODE = re.compile(r"\b([A-Z]{2}|[A-Z][0-9]|[0-9][A-Z])( ?)([0-9]{1,4})\b")
NOT_AIRLINES = frozenset({"AM", "PM"})

LATE_MINUTES = 15  # the usual airline line between on time and delayed
WINDOW_BEFORE = timedelta(hours=4)
BEFORE_EVERY = timedelta(minutes=30)
AIR_EVERY = timedelta(minutes=60)
GIVE_UP_AFTER = timedelta(hours=6)  # after the expected arrival with no final answer
RETRY_AFTER_ERROR = timedelta(minutes=30)

TERMINAL = frozenset({"Arrived", "Canceled", "CanceledUncertain", "Diverted"})
IN_AIR = frozenset({"Departed", "EnRoute", "Approaching"})
CANCELLED = frozenset({"Canceled", "CanceledUncertain"})


@dataclass(frozen=True)
class FlightInfo:
    """One flight on one day as the provider described it. Times are UTC; `*_local` is display."""

    code: str
    dep_date: date
    found: bool
    status: str = ""
    sched_dep: datetime | None = None
    revised_dep: datetime | None = None
    sched_arr: datetime | None = None
    revised_arr: datetime | None = None
    sched_dep_local: str = ""
    revised_dep_local: str = ""
    dep_airport: str = ""
    arr_airport: str = ""
    dep_gate: str = ""

    @property
    def departs(self) -> datetime | None:
        return self.revised_dep or self.sched_dep

    @property
    def arrives(self) -> datetime | None:
        return self.revised_arr or self.sched_arr

    @property
    def late_minutes(self) -> int:
        if self.sched_dep and self.revised_dep:
            return int((self.revised_dep - self.sched_dep).total_seconds() // 60)
        return 0


def find_codes(text: str) -> list[str]:
    """Flight codes in `text`, in order, without spaces or repeats ("AC 663" -> "AC663")."""
    out: list[str] = []
    for m in _CODE.finditer(text or ""):
        airline, space, number = m.groups()
        if airline in NOT_AIRLINES:
            continue
        if space and len(number) == 4 and 1970 <= int(number) <= 2100:
            continue  # "CA 2026 trip": a year, not a flight
        code = airline + number
        if code not in out:
            out.append(code)
    return out


def codes_for_day(o: Occurrence) -> list[str]:
    """The codes that belong to this day of an event (docs/21 section 2).

    A single-day event: all of them (a connection). A trip: the first code on its first day and the
    last code on its last day; a trip with one code shows it on the first day only.
    """
    codes = find_codes(f"{o.title} {o.location}")
    if not codes or o.day_count == 1:
        return codes
    if o.is_first:
        return codes[:1]
    if o.is_last and len(codes) > 1:
        return codes[-1:]
    return []


def word(info: FlightInfo | None) -> str:
    """The status word drawn after the title, or "" for nothing (docs/21 section 4)."""
    if info is None or not info.found:
        return ""
    s = info.status
    if s in CANCELLED:
        return "Cancelled"
    if s == "Diverted":
        return "Diverted"
    if s == "Arrived":
        return "Landed"
    if s in IN_AIR:
        return "Departed"
    if s == "Delayed" or info.late_minutes >= LATE_MINUTES:
        return "Delayed"
    return "On time" if info.sched_dep else ""


def pick(codes: list[str], infos: dict[str, FlightInfo]) -> tuple[str, FlightInfo | None]:
    """A connection shows its first flight until that one has landed, then the next."""
    for code in codes:
        info = infos.get(code)
        if info is None or not info.found or info.status != "Arrived":
            return code, info
    return codes[-1], infos.get(codes[-1])


def first_check_due(dep_date: date, today: date) -> bool:
    """A flight nobody has looked up yet is looked up from the day before it."""
    return dep_date - timedelta(days=1) <= today <= dep_date


def next_due(info: FlightInfo, now: datetime) -> datetime | None:
    """When to ask again after `info` arrived at `now`; None means never (docs/21 section 3)."""
    if not info.found or info.status in TERMINAL:
        return None
    dep, arr = info.departs, info.arrives
    if dep is None:
        return now + timedelta(hours=6)  # known flight but no times yet: try a few times a day
    if arr is not None and now > (info.sched_arr or arr) + GIVE_UP_AFTER:
        return None
    if info.status in IN_AIR or now >= dep:
        due = now + AIR_EVERY
        if arr is not None and now < arr < due:
            due = arr  # one look right at the expected arrival
        return due
    if now < dep - WINDOW_BEFORE:
        return dep - WINDOW_BEFORE
    return now + BEFORE_EVERY


def utc(value: datetime) -> datetime:
    return value if value.tzinfo else value.replace(tzinfo=UTC)
