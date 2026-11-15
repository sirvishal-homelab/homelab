"""Flight codes, the status word and the lookup schedule (docs/21 sections 2 to 4)."""

from datetime import UTC, date, datetime, timedelta

import pytest

from trmnl_calendar.domain import flights as F
from trmnl_calendar.domain.models import Occurrence

DAY = date(2026, 11, 16)
DEP = datetime(2026, 11, 16, 15, 25, tzinfo=UTC)  # 11:25 in Halifax
ARR = datetime(2026, 11, 16, 17, 9, tzinfo=UTC)


@pytest.mark.parametrize(
    "text, codes",
    [
        ("AC663", ["AC663"]),
        ("Flight AC 663", ["AC663"]),
        ("WS456", ["WS456"]),
        ("B6 789", ["B6789"]),
        ("U2 1234", ["U21234"]),
        ("Toronto AC604 / AC611", ["AC604", "AC611"]),
        ("AC663 then AC663 again", ["AC663"]),
        ("CA2026", ["CA2026"]),  # written without a space it is looked up
        ("Room B2", []),
        ("PM 5", []),
        ("Meet at 3 PM 10", []),
        ("Dentist 9 AM", []),
        ("CA 2026 trip", []),  # a year
        ("ac663", []),  # lower case is never matched
        ("Work (out)", []),
        ("Lunch at 12", []),
        ("", []),
    ],
)
def test_find_codes(text, codes):
    assert F.find_codes(text) == codes


def occ(title, d=DAY, day_number=1, day_count=1, location=""):
    return Occurrence(
        1, title, location, 1, d, "11:25", None, False, day_number=day_number, day_count=day_count
    )


def test_codes_for_day_single_day_keeps_all_codes_and_reads_location():
    assert F.codes_for_day(occ("Fly AC663 AC404")) == ["AC663", "AC404"]
    assert F.codes_for_day(occ("Fly to Montreal", location="AC663")) == ["AC663"]


def test_codes_for_day_on_a_trip_first_day_first_code_last_day_last_code():
    t = "Toronto AC604 / AC611"
    assert F.codes_for_day(occ(t, day_number=1, day_count=4)) == ["AC604"]
    assert F.codes_for_day(occ(t, day_number=2, day_count=4)) == []
    assert F.codes_for_day(occ(t, day_number=4, day_count=4)) == ["AC611"]


def test_a_trip_with_one_code_shows_it_on_the_first_day_only():
    assert F.codes_for_day(occ("Toronto AC604", day_number=1, day_count=3)) == ["AC604"]
    assert F.codes_for_day(occ("Toronto AC604", day_number=3, day_count=3)) == []


def info(status="Expected", sched=DEP, revised=None, **kw):
    return F.FlightInfo(
        "AC663", DAY, True, status, sched, revised, ARR, kw.pop("revised_arr", None), **kw
    )


@pytest.mark.parametrize(
    "flight, shown",
    [
        (None, ""),
        (F.FlightInfo("AC663", DAY, found=False), ""),
        (info("Expected"), "On time"),
        (info("Boarding"), "On time"),
        (info("Expected", revised=DEP + timedelta(minutes=14)), "On time"),
        (info("Expected", revised=DEP + timedelta(minutes=15)), "Delayed"),
        (info("Delayed"), "Delayed"),
        (info("Departed"), "Departed"),
        (info("EnRoute"), "Departed"),
        (info("Approaching"), "Departed"),
        (info("Arrived"), "Landed"),
        (info("Canceled"), "Cancelled"),
        (info("CanceledUncertain"), "Cancelled"),
        (info("Diverted"), "Diverted"),
        (info("Unknown", sched=None), ""),
        (info("Unknown"), "On time"),
    ],
)
def test_word(flight, shown):
    assert F.word(flight) == shown


def test_a_connection_shows_the_next_flight_once_the_first_has_landed():
    first, second = info("Arrived"), F.FlightInfo("AC404", DAY, True, "Expected", DEP)
    assert F.pick(["AC663", "AC404"], {"AC663": info()})[0] == "AC663"
    assert F.pick(["AC663", "AC404"], {"AC663": first, "AC404": second}) == ("AC404", second)
    assert F.pick(["AC663", "AC404"], {})[0] == "AC663"


def test_first_check_from_the_day_before():
    assert not F.first_check_due(DAY, DAY - timedelta(days=2))
    assert F.first_check_due(DAY, DAY - timedelta(days=1))
    assert F.first_check_due(DAY, DAY)
    assert not F.first_check_due(DAY, DAY + timedelta(days=1))


def test_schedule_follows_the_flight():
    early = DEP - timedelta(hours=20)
    assert F.next_due(info(), early) == DEP - timedelta(hours=4)
    inside = DEP - timedelta(hours=2)
    assert F.next_due(info(), inside) == inside + timedelta(minutes=30)
    flying = DEP + timedelta(minutes=10)
    assert F.next_due(info("EnRoute"), flying) == flying + timedelta(minutes=60)
    near = ARR - timedelta(minutes=20)
    assert F.next_due(info("Approaching"), near) == ARR  # one look at the expected arrival


def test_schedule_follows_a_delay():
    late = DEP + timedelta(hours=3)
    now = DEP - timedelta(hours=2)  # inside the old window, before the new one
    assert F.next_due(info("Delayed", revised=late), now) == late - timedelta(hours=4)


@pytest.mark.parametrize("status", ["Arrived", "Canceled", "CanceledUncertain", "Diverted"])
def test_final_answers_are_never_asked_again(status):
    assert F.next_due(info(status), DEP) is None


def test_not_found_is_never_asked_again_and_stale_flights_give_up():
    assert F.next_due(F.FlightInfo("AC663", DAY, found=False), DEP) is None
    assert F.next_due(info("EnRoute"), ARR + timedelta(hours=7)) is None


def test_a_known_flight_without_times_is_tried_a_few_times_a_day():
    now = DEP - timedelta(days=1)
    assert F.next_due(info("Unknown", sched=None), now) == now + timedelta(hours=6)
