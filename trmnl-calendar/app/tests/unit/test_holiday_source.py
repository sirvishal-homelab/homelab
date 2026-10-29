import json
from datetime import date

import pytest

from trmnl_calendar import holiday_source as S

CA_2026 = [
    {
        "date": "2026-01-01",
        "localName": "New Year's Day",
        "name": "New Year's Day",
        "countryCode": "CA",
        "global": True,
        "counties": None,
        "types": ["Public"],
    },
    {
        "date": "2026-02-16",
        "localName": "Heritage Day",
        "name": "Heritage Day",
        "global": False,
        "counties": ["CA-NS"],
        "types": ["Public"],
    },
    {
        "date": "2026-02-16",
        "localName": "Louis Riel Day",
        "name": "Louis Riel Day",
        "global": False,
        "counties": ["CA-MB"],
        "types": ["Public"],
    },
    {
        "date": "2026-11-11",
        "localName": "Remembrance Day",
        "name": "Remembrance Day",
        "global": False,
        "subdivisionCodes": ["CA-NS", "CA-AB"],
        "types": ["Bank", "Optional"],
    },
    {"date": "2026-12-24", "name": "Christmas Eve", "types": ["Observance"]},
    {"date": "not a date", "name": "Broken"},
    "junk",
]


class Fake:
    def __init__(self, replies):
        self.replies, self.calls = replies, []

    def __call__(self, url, timeout):
        self.calls.append(url)
        reply = self.replies.get(url.rsplit("/api/v3/", 1)[-1], (404, b""))
        status, payload = reply if isinstance(reply, tuple) else (200, reply)
        return status, payload if isinstance(payload, bytes) else json.dumps(payload).encode()


def source(replies, **kw):
    return S.HolidaySource("https://example.test/api/v3/", transport=Fake(replies), **kw)


def test_parse_skips_bad_entries_and_accepts_either_field_name():
    items = S.parse_holidays(CA_2026)
    assert [h.name for h in items] == [
        "New Year's Day",
        "Heritage Day",
        "Louis Riel Day",
        "Remembrance Day",
        "Christmas Eve",
    ]
    assert items[3].subdivisions == ("CA-AB", "CA-NS")  # subdivisionCodes, sorted
    assert items[1].subdivisions == ("CA-NS",)  # counties
    assert S.subdivision_codes(items) == ["CA-AB", "CA-MB", "CA-NS"]


def test_rows_for_a_subdivision_include_national_and_its_own_days():
    items = S.parse_holidays(CA_2026)
    names = {h.name: h.group for h in S.to_holidays(items, "ca-ns")}
    assert names == {
        "New Year's Day": "public",
        "Heritage Day": "public",
        "Remembrance Day": "closing",
        "Christmas Eve": "common",
    }
    assert {h.name for h in S.to_holidays(items, "")} == {"New Year's Day", "Christmas Eve"}
    row = next(h for h in S.to_holidays(items, "CA-NS") if h.name == "Heritage Day")
    assert row.key == "heritage_day" and row.date == date(2026, 2, 16)


def test_unusable_replies_are_errors_not_crashes():
    with pytest.raises(S.HolidaySourceError):
        S.parse_holidays({"error": "x"})
    with pytest.raises(S.HolidaySourceError):
        S.parse_holidays([{"nothing": 1}])
    assert S.parse_holidays([]) == []
    with pytest.raises(S.HolidaySourceError):
        S.parse_countries([])


def test_countries_sorted_and_cached():
    s = source(
        {
            "AvailableCountries": [
                {"countryCode": "DE", "name": "Germany"},
                {"countryCode": "CA", "name": "Canada"},
                {"countryCode": "toolong", "name": "x"},
            ]
        }
    )
    assert [c.code for c in s.countries()] == ["CA", "DE"]
    s.countries()
    assert len(s._get.calls) == 1


def test_holidays_cache_expires():
    now = [0.0]
    s = source({"PublicHolidays/2026/CA": CA_2026}, clock=lambda: now[0])
    s.holidays("ca", 2026)
    s.holidays("CA", 2026)
    assert len(s._get.calls) == 1
    now[0] = S.HOLIDAY_TTL + 1
    s.holidays("CA", 2026)
    assert len(s._get.calls) == 2


def test_lookup_returns_codes_and_rows():
    s = source({"PublicHolidays/2026/CA": CA_2026})
    codes, rows = s.lookup("CA", "CA-NS", [2026])
    assert codes == ["CA-AB", "CA-MB", "CA-NS"] and len(rows) == 4


@pytest.mark.parametrize(
    "reply, text",
    [
        ((404, b""), "no list"),
        ((500, b"oops"), "HTTP 500"),
        ((200, b"<html>"), "not JSON"),
        ((200, b"x" * (S.MAX_BYTES + 1)), "far more data"),
    ],
)
def test_service_failures_have_plain_messages(reply, text):
    s = source({"PublicHolidays/2026/CA": reply})
    with pytest.raises(S.HolidaySourceError, match=text):
        s.holidays("CA", 2026)


def test_bad_inputs():
    s = source({})
    with pytest.raises(S.HolidaySourceError):
        s.holidays("../x", 2026)
    with pytest.raises(ValueError):
        S.HolidaySource("file:///etc/passwd")
