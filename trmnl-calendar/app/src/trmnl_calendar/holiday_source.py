"""Look up public holidays for other countries from an online service (ADR-0011).

The built-in rule sets in `domain/holidays.py` stay the default. This module is the only place that
talks to the service. It is used only when someone clicks "Look up" on the holidays page, never on
a schedule, and it sends nothing but a country code and a year.

The service is Nager.Date (https://date.nager.at, MIT licence). Its documented v3 reply is a list of
objects with `date`, `name`, `localName`, `global`, `counties` and `types`. The maintainers say the
`counties` field will be renamed to `subdivisionCodes` and v3 will be supported until 2027-01-31, so
the parser accepts either name, and the address is configurable (`HOLIDAY_API_URL`).

Standard library only (urllib). Synchronous: the web layer runs it in a worker thread.
"""

from __future__ import annotations

import json
import re
import threading
import time
import urllib.error
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date

from trmnl_calendar.domain.holidays import Holiday

DEFAULT_URL = "https://date.nager.at/api/v3"
HOLIDAY_TTL = 60 * 60  # a year's list does not change; keep it an hour so a re-import is instant
COUNTRY_TTL = 24 * 60 * 60
MAX_BYTES = 1_000_000

# Reply -> the picker's groups (labels live in domain/holidays.GROUPS).
_GROUP_BY_TYPE = {"public": "public", "bank": "closing", "authorities": "closing"}


class HolidaySourceError(Exception):
    """The lookup failed. The message is written for the person using the page."""


@dataclass(frozen=True)
class Country:
    code: str
    name: str


@dataclass(frozen=True)
class OnlineHoliday:
    name: str
    local_name: str
    date: date
    types: tuple[str, ...]
    subdivisions: tuple[str, ...]  # empty: applies to the whole country


Transport = Callable[[str, float], tuple[int, bytes]]


def urllib_get(url: str, timeout: float) -> tuple[int, bytes]:
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:  # noqa: S310 (http/https checked)
            return r.status, r.read(MAX_BYTES + 1)
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read(MAX_BYTES + 1)
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        reason = getattr(exc, "reason", exc)
        raise HolidaySourceError(f"Could not reach the holiday service: {reason}") from exc


def slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_") or "holiday"


def _text(value: object) -> str:
    return value.strip() if isinstance(value, str) else ""


def _codes(value: object) -> tuple[str, ...]:
    if not isinstance(value, list):
        return ()
    return tuple(sorted({v.strip().upper() for v in value if isinstance(v, str) and v.strip()}))


def parse_holidays(payload: object) -> list[OnlineHoliday]:
    """Read the service's list; skip entries that are unusable rather than fail the whole page."""
    if not isinstance(payload, list):
        raise HolidaySourceError("The holiday service sent a reply this app does not understand.")
    out = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        try:
            day = date.fromisoformat(_text(item.get("date"))[:10])
        except ValueError:
            continue
        name = _text(item.get("name")) or _text(item.get("localName"))
        if not name:
            continue
        subs = _codes(
            item.get("subdivisionCodes") or item.get("counties") or item.get("subdivisions")
        )
        types = tuple(t for t in (item.get("types") or ["Public"]) if isinstance(t, str))
        out.append(
            OnlineHoliday(name, _text(item.get("localName")), day, types or ("Public",), subs)
        )
    if payload and not out:
        raise HolidaySourceError("The holiday service sent a reply this app does not understand.")
    return out


def parse_countries(payload: object) -> list[Country]:
    if not isinstance(payload, list):
        raise HolidaySourceError("The holiday service sent a reply this app does not understand.")
    out = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        code, name = _text(item.get("countryCode")).upper(), _text(item.get("name"))
        if re.fullmatch(r"[A-Z]{2}", code) and name:
            out.append(Country(code, name))
    if not out:
        raise HolidaySourceError("The holiday service sent no countries.")
    return sorted(out, key=lambda c: c.name.casefold())


def subdivision_codes(items: list[OnlineHoliday]) -> list[str]:
    return sorted({s for h in items for s in h.subdivisions})


def to_holidays(items: list[OnlineHoliday], subdivision: str = "") -> list[Holiday]:
    """The picker's rows: nationwide days, plus the chosen subdivision's own days."""
    sub = subdivision.strip().upper()
    out: dict[tuple[str, date], Holiday] = {}
    for h in items:
        if h.subdivisions and sub not in h.subdivisions:
            continue
        kinds = {t.lower() for t in h.types}
        group = next(
            (_GROUP_BY_TYPE[k] for k in ("public", "bank", "authorities") if k in kinds), "common"
        )
        key = slug(h.name)
        out.setdefault((key, h.date), Holiday(key, h.name, h.date, group))
    return sorted(out.values(), key=lambda x: (x.date, x.name))


class HolidaySource:
    def __init__(
        self,
        base_url: str = DEFAULT_URL,
        *,
        transport: Transport | None = None,
        timeout: float = 8.0,
        clock: Callable[[], float] = time.monotonic,
    ):
        self.base = base_url.rstrip("/")
        if not self.base.lower().startswith(("http://", "https://")):
            raise ValueError("HOLIDAY_API_URL must start with http:// or https://")
        self._get = transport or urllib_get
        self._timeout = timeout
        self._clock = clock
        self._cache: dict[str, tuple[float, object]] = {}
        self._lock = threading.Lock()

    def _json(self, path: str, ttl: float) -> object:
        now = self._clock()
        with self._lock:
            hit = self._cache.get(path)
            if hit and now - hit[0] < ttl:
                return hit[1]
        status, body = self._get(f"{self.base}/{path}", self._timeout)
        if status == 404:
            raise HolidaySourceError("The holiday service has no list for that country and year.")
        if status != 200:
            raise HolidaySourceError(f"The holiday service answered with an error (HTTP {status}).")
        if len(body) > MAX_BYTES:
            raise HolidaySourceError("The holiday service sent far more data than expected.")
        try:
            data = json.loads(body)
        except ValueError:
            raise HolidaySourceError("The holiday service sent a reply that is not JSON.") from None
        with self._lock:
            self._cache[path] = (now, data)
        return data

    def countries(self) -> list[Country]:
        return parse_countries(self._json("AvailableCountries", COUNTRY_TTL))

    def holidays(self, country: str, year: int) -> list[OnlineHoliday]:
        code = country.strip().upper()
        if not re.fullmatch(r"[A-Z]{2}", code):
            raise HolidaySourceError("Country must be a two-letter code such as CA.")
        return parse_holidays(self._json(f"PublicHolidays/{int(year)}/{code}", HOLIDAY_TTL))

    def lookup(
        self, country: str, subdivision: str, years: list[int]
    ) -> tuple[list[str], list[Holiday]]:
        """(subdivision codes seen in these years, the holiday rows for the chosen subdivision)."""
        items = [h for y in years for h in self.holidays(country, y)]
        return subdivision_codes(items), to_holidays(items, subdivision)
