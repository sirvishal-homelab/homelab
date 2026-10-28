"""Typed domain objects. Pure: no database, no clock, no I/O."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

WEEKDAY_CODES = ("MO", "TU", "WE", "TH", "FR", "SA", "SU")  # index == date.weekday()
WEEKDAY_NAMES = (
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
)
FREQS = ("none", "daily", "weekly", "monthly", "yearly")


@dataclass(frozen=True)
class Event:
    id: int
    title: str
    section_id: int
    start_date: date
    end_date: date | None = None  # last day of a multi-day event (trip)
    start_time: str | None = None  # "HH:MM"; None means all-day
    end_time: str | None = None
    location: str = ""
    freq: str = "none"
    interval: int = 1
    by_weekday: tuple[str, ...] = ()
    month_mode: str | None = None  # "date" | "weekday"
    month_nth: int | None = None  # 1..4 or -1
    until: date | None = None
    uid: str = ""
    rev: int = 1
    people: tuple[int, ...] = ()  # person ids, owner first; empty means everyone
    owner_id: int | None = None

    @property
    def duration_days(self) -> int:
        return (self.end_date - self.start_date).days if self.end_date else 0


@dataclass(frozen=True)
class Occurrence:
    """One event on one day. A trip yields one Occurrence per day it covers."""

    event_id: int
    title: str
    location: str
    section_id: int
    date: date
    start_time: str | None  # shown time for this day, None if none to show
    end_time: str | None
    all_day: bool
    day_number: int = 1
    day_count: int = 1
    series_start: date | None = None  # start date of the occurrence this day belongs to
    people: tuple[int, ...] = field(default_factory=tuple)  # owner first
    owner_id: int | None = None

    @property
    def is_first(self) -> bool:
        return self.day_number == 1

    @property
    def is_last(self) -> bool:
        return self.day_number == self.day_count

    @property
    def is_multi_day(self) -> bool:
        return self.day_count > 1


@dataclass(frozen=True)
class Section:
    id: int
    name: str
    position: int
    weight: int = 1
    visible: bool = True
    placement: str = "lane"  # "lane" | "banner"


@dataclass(frozen=True)
class CalendarInfo:
    id: int
    name: str
    timezone: str
    week_start: str = "sun"  # "sun" | "mon"
    time_format: str = "12"  # "12" | "24"


@dataclass(frozen=True)
class Person:
    id: int
    name: str
    label: str  # one or two characters, drawn on the screen
    position: int = 0
    active: bool = True
