"""The AeroDataBox client, against made-up replies (no test calls the live service)."""

import json
from datetime import UTC, date, datetime

import pytest

from trmnl_calendar import flight_source as S

DAY = date(2026, 11, 16)

# The documented v1 shape, with made-up values.
LEG = {
    "number": "AC 663",
    "status": "Expected",
    "departure": {
        "airport": {"icao": "CYHZ", "iata": "YHZ", "name": "Halifax"},
        "scheduledTime": {"utc": "2026-11-16 15:25Z", "local": "2026-11-16 11:25-04:00"},
        "revisedTime": {"utc": "2026-11-16 15:55Z", "local": "2026-11-16 11:55-04:00"},
        "terminal": "",
        "gate": "23",
    },
    "arrival": {
        "airport": {"icao": "CYUL", "iata": "YUL", "name": "Montreal"},
        "scheduledTime": {"utc": "2026-11-16 17:09Z", "local": "2026-11-16 12:09-05:00"},
    },
}


class Fake:
    def __init__(self, status=200, body=b""):
        self.status, self.body, self.calls = status, body, []

    def __call__(self, url, headers, timeout):
        self.calls.append((url, headers))
        return self.status, self.body


def src(status=200, body=b"", key="secret-key"):
    fake = Fake(status, body if isinstance(body, bytes) else json.dumps(body).encode())
    return S.FlightSource(key, "https://aerodatabox.p.rapidapi.com/", transport=fake), fake


def test_lookup_sends_code_date_and_rapidapi_headers_only():
    s, fake = src(body=[LEG])
    s.lookup("AC663", DAY)
    url, headers = fake.calls[0]
    assert url.startswith("https://aerodatabox.p.rapidapi.com/flights/number/AC663/2026-11-16?")
    assert headers["X-RapidAPI-Key"] == "secret-key"
    assert headers["X-RapidAPI-Host"] == "aerodatabox.p.rapidapi.com"


def test_parse_reads_times_airports_gate_and_status():
    s, _ = src(body=[LEG])
    i = s.lookup("AC663", DAY)
    assert i.found and i.status == "Expected"
    assert i.sched_dep == datetime(2026, 11, 16, 15, 25, tzinfo=UTC)
    assert i.revised_dep == datetime(2026, 11, 16, 15, 55, tzinfo=UTC)
    assert i.sched_arr == datetime(2026, 11, 16, 17, 9, tzinfo=UTC) and i.revised_arr is None
    assert (i.sched_dep_local, i.revised_dep_local) == ("11:25", "11:55")
    assert (i.dep_airport, i.arr_airport, i.dep_gate) == ("YHZ", "YUL", "23")
    assert i.late_minutes == 30


def test_parse_accepts_the_older_flat_time_fields():
    old = {
        "status": "Delayed",
        "departure": {
            "airport": {"iata": "YHZ"},
            "scheduledTimeUtc": "2026-11-16 15:25Z",
            "scheduledTimeLocal": "2026-11-16 11:25-04:00",
        },
        "arrival": {"airport": {"iata": "YUL"}},
    }
    i = S.parse("AC663", DAY, [old])
    assert i.sched_dep == datetime(2026, 11, 16, 15, 25, tzinfo=UTC)
    assert i.sched_dep_local == "11:25" and i.status == "Delayed"


def test_parse_picks_the_leg_departing_that_day():
    other = json.loads(json.dumps(LEG))
    other["departure"]["scheduledTime"]["local"] = "2026-11-15 11:25-04:00"
    other["status"] = "Arrived"
    assert S.parse("AC663", DAY, [other, LEG]).status == "Expected"


@pytest.mark.parametrize("status, body", [(204, b""), (404, b"{}"), (200, b""), (200, b"[]")])
def test_no_such_flight_is_not_found_not_an_error(status, body):
    s, _ = src(status, body)
    assert s.lookup("AC663", DAY).found is False


@pytest.mark.parametrize(
    "status, body, text",
    [
        (401, b"", "refused the key"),
        (403, b"", "refused the key"),
        (429, b"", "allowance"),
        (500, b"", "HTTP 500"),
        (200, b"<html>", "not JSON"),
        (200, b'{"a": 1}', "does not understand"),
    ],
)
def test_failures_have_plain_messages_and_never_show_the_key(status, body, text):
    s, _ = src(status, body)
    with pytest.raises(S.FlightSourceError, match=text) as exc:
        s.lookup("AC663", DAY)
    assert "secret-key" not in str(exc.value)


def test_no_key_and_plain_http_are_refused_and_repr_hides_the_key():
    s, fake = src(key="")
    with pytest.raises(S.FlightSourceError, match="FLIGHT_API_KEY"):
        s.lookup("AC663", DAY)
    assert fake.calls == []
    with pytest.raises(ValueError):
        S.FlightSource("k", "http://aerodatabox.p.rapidapi.com")
    assert "secret" not in repr(src()[0])
