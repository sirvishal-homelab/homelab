"""Layout interface pieces shared by all layouts (docs/04 section 5). Pure data and helpers."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from trmnl_calendar.domain.models import Occurrence, Person, Section

DAY_ABBR = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")  # index == date.weekday()
MONTH_ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
ALL_DAYS = 127  # bit i set = weekday i shown (0 = Monday)

# Options the editor knows about; each layout lists the subset it uses (its OPTIONS).
OPTION_LABELS = {
    "show_times": "Show times",
    "show_badges": "Show who it is for (initials)",
    "show_locations": "Show locations",
    "show_banners": "Show banner categories (holidays and similar)",
    "show_hidden_marker": "Note events on hidden days",
    "hide_empty_lanes": "Hide categories with nothing this week",
}
OPTION_DEFAULTS = {
    "show_times": True,
    "show_badges": True,
    "show_locations": False,
    "show_banners": True,
    "show_hidden_marker": True,
    "hide_empty_lanes": False,
}


@dataclass(frozen=True)
class PlanContext:
    today: date
    sections: tuple[Section, ...]  # already filtered to visible, in display order
    occurrences: tuple[Occurrence, ...]
    week_start: str = "sun"
    time_format: str = "12"
    weekdays: int = ALL_DAYS
    width: int = 800
    height: int = 480
    bit_depth: int = 2
    options: dict = field(default_factory=dict)
    people: dict[int, Person] = field(default_factory=dict)  # for badges

    def opt(self, key: str) -> bool:
        return bool(self.options.get(key, OPTION_DEFAULTS[key]))


def fmt_time(hhmm: str, time_format: str) -> str:
    h, m = int(hhmm[:2]), int(hhmm[3:5])
    if time_format == "24":
        return f"{h:02d}:{m:02d}"
    suffix = "a" if h < 12 else "p"
    h12 = h % 12 or 12
    return f"{h12}{suffix}" if m == 0 else f"{h12}:{m:02d}{suffix}"


def fmt_range(first: date, last: date) -> str:
    """'Nov 3–6' or 'Oct 30 – Nov 2'; with the year only when it differs between the ends."""
    a = f"{MONTH_ABBR[first.month - 1]} {first.day}"
    if first == last:
        return a
    if first.month == last.month and first.year == last.year:
        return f"{a}–{last.day}"
    return f"{a} – {MONTH_ABBR[last.month - 1]} {last.day}"


def title_range(first: date, last: date) -> str:
    """Heading for a week: 'Oct 25 – 31, 2026'."""
    if first.year != last.year:
        return (
            f"{MONTH_ABBR[first.month - 1]} {first.day}, {first.year} – "
            f"{MONTH_ABBR[last.month - 1]} {last.day}, {last.year}"
        )
    return f"{fmt_range(first, last).replace('–', ' – ', 1) if first.month == last.month else fmt_range(first, last)}, {last.year}"  # noqa: E501
