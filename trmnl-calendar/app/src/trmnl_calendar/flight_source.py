"""Look up one flight's status for one day from AeroDataBox through RapidAPI (ADR-0015, docs/21).

The only module that talks to the flight service. It sends the flight code and the date, nothing
else. The key goes in the `X-RapidAPI-Key` header and is never logged or put in a message.

Endpoint: GET {base}/flights/number/{code}/{YYYY-MM-DD} ("Flight status (specific date)", a tier-2
endpoint: 2 units per call on the Basic plan). The reply is a list of legs. Recent versions give
times as {"utc": "2026-11-16 15:25Z", "local": "2026-11-16 11:25-04:00"}; older ones as
`scheduledTimeUtc` / `scheduledTimeLocal`. Both are read. A day with no such flight answers 204
or 404. This parser was written from the documentation, so the first live reply is its check
(docs/08 R-15).

Standard library only (urllib). Synchronous: callers run it in a worker thread.
"""

from __future__ import annotations

import json
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable
from datetime import UTC, date, datetime

from trmnl_calendar.domain.flights import FlightInfo

DEFAULT_URL = "https://aerodatabox.p.rapidapi.com"
MAX_BYTES = 500_000


class FlightSourceError(Exception):
    """The lookup failed. The message is written for the person using the page."""


Transport = Callable[[str, dict[str, str], float], tuple[int, bytes]]


def urllib_get(url: str, headers: dict[str, str], timeout: float) -> tuple[int, bytes]:
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:  # noqa: S310 (https checked)
            return r.status, r.read(MAX_BYTES + 1)
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read(MAX_BYTES + 1)
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        reason = getattr(exc, "reason", exc)
        raise FlightSourceError(f"Could not reach the flight service: {reason}") from exc


def _time(value: object) -> datetime | None:
    """'2026-11-16 15:25Z', '2026-11-16T15:25:00+00:00' or '2026-11-16 11:25-04:00' -> UTC."""
    if not isinstance(value, str) or not value.strip():
        return None
    text = value.strip().replace(" ", "T")
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        t = datetime.fromisoformat(text)
    except ValueError:
        return None
    if t.tzinfo is None:
        t = t.replace(tzinfo=UTC)
    return t.astimezone(UTC)


def _local_text(value: object) -> str:
    """'2026-11-16 11:25-04:00' -> '11:25' (the airport's own clock, for display)."""
    if not isinstance(value, str):
        return ""
    text = value.strip().replace(" ", "T")
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).strftime("%H:%M")
    except ValueError:
        return ""


def _pair(side: dict, key: str) -> tuple[datetime | None, str]:
    """(UTC time, local display text) for scheduledTime / revisedTime in either reply shape."""
    node = side.get(key)
    if isinstance(node, dict):
        utc_t = _time(node.get("utc")) or _time(node.get("local"))
        return utc_t, _local_text(node.get("local"))
    cap = key[0].upper() + key[1:]
    utc_t = _time(side.get(f"{key}Utc")) or _time(side.get(f"{cap}Utc"))
    local = side.get(f"{key}Local") or side.get(f"{cap}Local")
    return utc_t or _time(local), _local_text(local)


def _airport(side: dict) -> str:
    ap = side.get("airport")
    if not isinstance(ap, dict):
        return ""
    for k in ("iata", "icao", "shortName", "name"):
        v = ap.get(k)
        if isinstance(v, str) and v.strip():
            return v.strip()
    return ""


def _text(value: object) -> str:
    return value.strip() if isinstance(value, str) else ""


def parse(code: str, day: date, payload: object) -> FlightInfo:
    """The leg that departs on `day` (local), else the first leg."""
    if not isinstance(payload, list):
        raise FlightSourceError("The flight service sent a reply this app does not understand.")
    legs = [leg for leg in payload if isinstance(leg, dict)]
    if not legs:
        return FlightInfo(code, day, found=False)

    def dep_side(leg: dict) -> dict:
        d = leg.get("departure")
        return d if isinstance(d, dict) else {}

    def local_day(leg: dict) -> str:
        node = dep_side(leg).get("scheduledTime")
        local = (
            node.get("local") if isinstance(node, dict) else dep_side(leg).get("scheduledTimeLocal")
        )
        return _text(local)[:10]

    leg = next((x for x in legs if local_day(x) == day.isoformat()), legs[0])
    dep = dep_side(leg)
    arr = leg.get("arrival") if isinstance(leg.get("arrival"), dict) else {}
    sched_dep, sched_dep_local = _pair(dep, "scheduledTime")
    revised_dep, revised_dep_local = _pair(dep, "revisedTime")
    sched_arr, _ = _pair(arr, "scheduledTime")
    revised_arr, _ = _pair(arr, "revisedTime")
    return FlightInfo(
        code=code,
        dep_date=day,
        found=True,
        status=_text(leg.get("status")) or "Unknown",
        sched_dep=sched_dep,
        revised_dep=revised_dep,
        sched_arr=sched_arr,
        revised_arr=revised_arr,
        sched_dep_local=sched_dep_local,
        revised_dep_local=revised_dep_local,
        dep_airport=_airport(dep),
        arr_airport=_airport(arr),
        dep_gate=_text(dep.get("gate")),
    )


class FlightSource:
    def __init__(
        self,
        api_key: str,
        base_url: str = DEFAULT_URL,
        *,
        transport: Transport | None = None,
        timeout: float = 10.0,
    ):
        self.base = (base_url or DEFAULT_URL).rstrip("/")
        if not self.base.lower().startswith("https://"):
            raise ValueError("FLIGHT_API_URL must start with https:// (the key is sent with it)")
        self._key = api_key
        self._host = urllib.parse.urlparse(self.base).hostname or ""
        self._get = transport or urllib_get
        self._timeout = timeout

    def __repr__(self) -> str:  # never show the key
        return f"FlightSource({self.base!r})"

    def lookup(self, code: str, day: date) -> FlightInfo:
        if not self._key:
            raise FlightSourceError("No flight API key is set (FLIGHT_API_KEY in the .env file).")
        path = f"flights/number/{urllib.parse.quote(code)}/{day.isoformat()}"
        query = "withAircraftImage=false&withLocation=false&withFlightPlan=false"
        headers = {
            "X-RapidAPI-Key": self._key,
            "X-RapidAPI-Host": self._host,
            "Accept": "application/json",
        }
        status, body = self._get(f"{self.base}/{path}?{query}", headers, self._timeout)
        if status in (204, 404):
            return FlightInfo(code, day, found=False)
        if status in (401, 403):
            raise FlightSourceError(
                "The flight service refused the key. Check FLIGHT_API_KEY and that you are "
                "subscribed to AeroDataBox on RapidAPI."
            )
        if status == 429:
            raise FlightSourceError("The flight service says this month's allowance is used up.")
        if status != 200:
            raise FlightSourceError(f"The flight service answered with an error (HTTP {status}).")
        if len(body) > MAX_BYTES:
            raise FlightSourceError("The flight service sent far more data than expected.")
        if not body.strip():
            return FlightInfo(code, day, found=False)
        try:
            data = json.loads(body)
        except ValueError:
            raise FlightSourceError("The flight service sent a reply that is not JSON.") from None
        return parse(code, day, data)
