"""Holiday lists by region, computed from rules (no network, no extra library). docs/10 slice 4.

Nova Scotia follows the province's own chart (novascotia.ca, Labour Standards): six paid holidays,
three retail closing days, and Remembrance Day under its own Act. Victoria Day, Easter Monday and
Natal Day are commonly observed but not statutory there. Dates are the calendar dates; a holiday
that falls on a weekend is not moved.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, timedelta

GROUPS = (
    ("public", "Public holidays"),
    ("closing", "Retail closing days and observances"),
    ("common", "Commonly observed, not statutory here"),
    ("days", "Other days"),
)
CHECKED_BY_DEFAULT = ("public", "closing")


@dataclass(frozen=True)
class Holiday:
    key: str
    name: str
    date: date
    group: str


def easter(year: int) -> date:
    """Western Easter Sunday (anonymous Gregorian algorithm)."""
    a, b, c = year % 19, year // 100, year % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    el = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * el) // 451
    month = (h + el - 7 * m + 114) // 31
    day = (h + el - 7 * m + 114) % 31 + 1
    return date(year, month, day)


def nth_weekday(year: int, month: int, weekday: int, n: int) -> date:
    """The nth (1-based) given weekday (Monday = 0) of a month."""
    first = date(year, month, 1)
    return first + timedelta(days=(weekday - first.weekday()) % 7 + 7 * (n - 1))


def _victoria_day(y: int) -> date:  # the Monday before May 25
    d = date(y, 5, 24)
    return d - timedelta(days=d.weekday())


# key -> (name, date rule)
RULES: dict[str, tuple[str, Callable[[int], date]]] = {
    "new_year": ("New Year's Day", lambda y: date(y, 1, 1)),
    "valentines": ("Valentine's Day", lambda y: date(y, 2, 14)),
    "heritage_day": ("Heritage Day", lambda y: nth_weekday(y, 2, 0, 3)),
    "dst_forward": ("Clocks forward", lambda y: nth_weekday(y, 3, 6, 2)),
    "good_friday": ("Good Friday", lambda y: easter(y) - timedelta(days=2)),
    "easter_sunday": ("Easter Sunday", easter),
    "easter_monday": ("Easter Monday", lambda y: easter(y) + timedelta(days=1)),
    "mothers_day": ("Mother's Day", lambda y: nth_weekday(y, 5, 6, 2)),
    "victoria_day": ("Victoria Day", _victoria_day),
    "fathers_day": ("Father's Day", lambda y: nth_weekday(y, 6, 6, 3)),
    "canada_day": ("Canada Day", lambda y: date(y, 7, 1)),
    "natal_day": ("Natal Day (Halifax area)", lambda y: nth_weekday(y, 8, 0, 1)),
    "labour_day": ("Labour Day", lambda y: nth_weekday(y, 9, 0, 1)),
    "truth_reconciliation": (
        "National Day for Truth and Reconciliation",
        lambda y: date(y, 9, 30),
    ),
    "thanksgiving": ("Thanksgiving", lambda y: nth_weekday(y, 10, 0, 2)),
    "halloween": ("Halloween", lambda y: date(y, 10, 31)),
    "dst_back": ("Clocks back", lambda y: nth_weekday(y, 11, 6, 1)),
    "remembrance_day": ("Remembrance Day", lambda y: date(y, 11, 11)),
    "christmas": ("Christmas Day", lambda y: date(y, 12, 25)),
    "boxing_day": ("Boxing Day", lambda y: date(y, 12, 26)),
}

_DAYS = {
    "valentines": "days",
    "mothers_day": "days",
    "fathers_day": "days",
    "halloween": "days",
    "truth_reconciliation": "days",
    "dst_forward": "days",
    "dst_back": "days",
}


def _groups(**by_group: tuple[str, ...]) -> dict[str, str]:
    out = {key: group for group, keys in by_group.items() for key in keys}
    return {**out, **_DAYS}


_CA_PUBLIC = (
    "new_year",
    "good_friday",
    "victoria_day",
    "canada_day",
    "labour_day",
    "thanksgiving",
    "christmas",
)
_CA_COMMON = ("easter_sunday", "easter_monday", "remembrance_day", "boxing_day")
_NS_PUBLIC = ("new_year", "heritage_day", "good_friday", "canada_day", "labour_day", "christmas")
_NS_CLOSING = ("easter_sunday", "thanksgiving", "remembrance_day", "boxing_day")
_NS_COMMON = ("victoria_day", "easter_monday", "natal_day")

REGIONS: dict[str, dict] = {
    "CA": {
        "label": "Canada (national)",
        "groups": _groups(public=_CA_PUBLIC, common=_CA_COMMON),
        "note": "Provinces differ. Pick your province if it is listed.",
    },
    "CA-NS": {
        "label": "Nova Scotia, Canada",
        "groups": _groups(public=_NS_PUBLIC, closing=_NS_CLOSING, common=_NS_COMMON),
        "note": "From the Nova Scotia Labour Standards holiday chart. "
        "Natal Day is a Halifax-area holiday.",
    },
}


def region_choices() -> list[tuple[str, str]]:
    return [(k, v["label"]) for k, v in REGIONS.items()]


def holidays_for(region: str, year: int) -> list[Holiday]:
    if region not in REGIONS:
        raise KeyError(region)
    out = []
    for key, group in REGIONS[region]["groups"].items():
        name, rule = RULES[key]
        out.append(Holiday(key, name, rule(year), group))
    return sorted(out, key=lambda h: (h.date, h.name))
