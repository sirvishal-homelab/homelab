"""Flight status end to end without the network: storage, schedule, cap, purge, backups, drawing."""

import sqlite3
from dataclasses import replace
from datetime import UTC, date, datetime, timedelta

import pytest

from trmnl_calendar import flight_service as svc
from trmnl_calendar.domain.flights import FlightInfo
from trmnl_calendar.errors import ValidationFailed
from trmnl_calendar.flight_source import FlightSourceError
from trmnl_calendar.storage import backup, db, migrate
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as events_repo
from trmnl_calendar.storage.repositories import flights as repo
from trmnl_calendar.storage.repositories import views as views_repo
from trmnl_calendar.web import render_service

DAY = date(2026, 11, 16)
DEP = datetime(2026, 11, 16, 15, 25, tzinfo=UTC)  # 11:25 in Halifax
ARR = datetime(2026, 11, 16, 17, 9, tzinfo=UTC)
EVENING_BEFORE = datetime(2026, 11, 15, 23, 0, tzinfo=UTC)  # 19:00 in Halifax, the day before


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


class FakeSource:
    def __init__(self, answers=None):
        self.answers = answers or {}
        self.calls = []

    def lookup(self, code, day, home=""):
        self.calls.append((code, day))
        a = self.answers.get(code)
        if isinstance(a, Exception):
            raise a
        return replace(a, code=code, dep_date=day) if a else FlightInfo(code, day, found=False)


def flight(status="Expected", revised=None, revised_local=""):
    return FlightInfo(
        "AC663", DAY, True, status, DEP, revised, ARR, None, "11:25", revised_local, "YHZ", "YUL"
    )


def add(conn, title="AC663", start=DAY, **kw):
    return events_repo.create_event(
        conn,
        {
            "title": title,
            "section_id": 1,
            "start_date": start.isoformat(),
            "start_time": "11:25",
            **kw,
        },
    )


def refresher(answers=None):
    src = FakeSource(answers)
    return svc.FlightRefresher(src), src


def test_migration_adds_the_tables(conn):
    names = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"flight_status", "flight_calls"} <= names


def test_first_lookup_happens_the_day_before_and_not_earlier(conn):
    add(conn)
    r, src = refresher({"AC663": flight()})
    assert r.refresh(conn, EVENING_BEFORE - timedelta(days=1)) is False
    assert src.calls == []
    assert r.refresh(conn, EVENING_BEFORE) is True  # nothing shown -> "On time"
    assert src.calls == [("AC663", DAY)]


def test_after_the_first_lookup_it_waits_for_the_window(conn):
    add(conn)
    r, src = refresher({"AC663": flight()})
    r.refresh(conn, EVENING_BEFORE)
    r.refresh(conn, EVENING_BEFORE + timedelta(hours=2))
    assert len(src.calls) == 1
    r.refresh(conn, DEP - timedelta(hours=4))  # the 4-hour window opens
    assert len(src.calls) == 2


def test_status_word_is_added_to_the_title_everywhere_occurrences_are_built(conn):
    add(conn)
    add(conn, title="Dentist")
    r, _ = refresher({"AC663": flight("Delayed")})
    r.refresh(conn, EVENING_BEFORE)
    occ = svc.annotate(conn, _occ(conn))
    titles = sorted(o.title for o in occ)
    assert titles == ["AC663 (Delayed)", "Dentist"]
    assert [o.flight for o in occ if o.flight] == ["Delayed"]


def _occ(conn):
    from trmnl_calendar.domain.recurrence import occurrences

    evs = events_repo.list_events(conn, DAY, DAY)
    return occurrences(evs, None, DAY, DAY)


def test_a_change_of_word_is_reported_so_the_screen_is_pushed(conn):
    add(conn)
    r, src = refresher({"AC663": flight()})
    r.refresh(conn, EVENING_BEFORE)
    now = DEP - timedelta(hours=3)
    src.answers["AC663"] = flight()  # same word: no push
    assert r.refresh(conn, now) is False
    src.answers["AC663"] = flight("Delayed")
    assert r.refresh(conn, now + timedelta(minutes=30)) is True


def test_not_found_costs_one_call_and_shows_nothing(conn):
    add(conn, title="Meeting room AB 12")
    r, src = refresher()
    r.refresh(conn, EVENING_BEFORE)
    r.refresh(conn, DEP)
    assert src.calls == [("AB12", DAY)]
    assert [o.title for o in svc.annotate(conn, _occ(conn))] == ["Meeting room AB 12"]


def test_a_failure_keeps_the_last_answer_and_retries_later(conn):
    add(conn)
    r, src = refresher({"AC663": flight("Delayed")})
    r.refresh(conn, EVENING_BEFORE)
    src.answers["AC663"] = FlightSourceError("Could not reach the flight service: timeout")
    now = DEP - timedelta(hours=3)
    assert r.refresh(conn, now) is False
    shown = svc.stored(conn, "AC663", DAY)
    assert shown.word == "Delayed" and "timeout" in shown.error
    r.refresh(conn, now + timedelta(minutes=10))
    assert len(src.calls) == 2  # waits 30 minutes before trying again


def test_monthly_cap_stops_lookups_and_counts_calls(conn):
    add(conn)
    add(conn, title="AC404")
    cal.update_calendar(conn, {"flight_monthly_limit": 1})
    r, src = refresher({"AC663": flight(), "AC404": flight()})
    r.refresh(conn, EVENING_BEFORE)
    assert len(src.calls) == 1 and r.capped
    assert repo.calls_this_month(conn, EVENING_BEFORE) == 1


def test_check_now_ignores_the_schedule_but_not_the_cap(conn):
    ev = add(conn, title="AC663 then AC404")
    r, src = refresher({"AC663": flight(), "AC404": flight()})
    results = r.check_now(conn, ev, EVENING_BEFORE - timedelta(days=5))
    assert [(x.code, x.word) for x in results] == [("AC663", "On time"), ("AC404", "On time")]
    cal.update_calendar(conn, {"flight_monthly_limit": 2})
    r.check_now(conn, ev, EVENING_BEFORE - timedelta(days=5))
    assert len(src.calls) == 2


def test_trip_uses_first_code_on_first_day_and_last_on_last(conn):
    ev = add(conn, title="Toronto AC604 / AC611", end_date="2026-11-19")
    assert svc.event_flights(ev) == [("AC604", DAY), ("AC611", date(2026, 11, 19))]


def test_old_answers_are_deleted_after_seven_days(conn):
    add(conn)
    r, _ = refresher({"AC663": flight()})
    r.refresh(conn, EVENING_BEFORE)
    assert repo.purge(conn, EVENING_BEFORE + timedelta(days=6)) == 0
    assert repo.purge(conn, EVENING_BEFORE + timedelta(days=8)) == 1


def test_backups_leave_flight_data_out(conn, settings):
    add(conn)
    refresher({"AC663": flight()})[0].refresh(conn, EVENING_BEFORE)
    path = backup.make_backup(settings.db_path, settings.backup_dir)
    copy = sqlite3.connect(path)
    try:
        assert copy.execute("SELECT COUNT(*) FROM flight_status").fetchone()[0] == 0
        assert copy.execute("SELECT COUNT(*) FROM flight_calls").fetchone()[0] == 1
        assert copy.execute("SELECT COUNT(*) FROM events").fetchone()[0] == 1
    finally:
        copy.close()


def test_the_screen_image_changes_when_the_status_arrives(conn):
    add(conn)
    view = views_repo.active_view(conn)
    before = render_service.render_view(conn, view, EVENING_BEFORE).hash
    refresher({"AC663": flight("Delayed")})[0].refresh(conn, EVENING_BEFORE)
    after = render_service.render_view(conn, view, EVENING_BEFORE)
    assert after.hash != before
    texts = [getattr(i, "text", "") for i in after.scene.items]
    assert any("(Delayed)" in t for t in texts)


def test_delayed_shows_the_new_time_in_the_calendar_style_and_a_new_time_is_pushed(conn):
    add(conn)
    late = flight("Delayed", DEP + timedelta(hours=2), "13:25")
    r, src = refresher({"AC663": late})
    r.refresh(conn, EVENING_BEFORE)
    assert [o.title for o in svc.annotate(conn, _occ(conn))] == ["AC663 (Delayed 1:25p)"]
    cal.update_calendar(conn, {"time_format": "24"})
    assert [o.title for o in svc.annotate(conn, _occ(conn))] == ["AC663 (Delayed 13:25)"]
    src.answers["AC663"] = flight("Delayed", DEP + timedelta(hours=3), "14:25")
    assert r.refresh(conn, DEP - timedelta(hours=1)) is True  # same word, later time: push


def add_untimed(conn, title="AC663", **kw):
    return events_repo.create_event(
        conn, {"title": title, "section_id": 1, "start_date": DAY.isoformat(), **kw}
    )


def test_a_flight_event_with_no_time_gets_the_planned_departure(conn):
    ev = add_untimed(conn)
    r, _ = refresher({"AC663": flight("Delayed", DEP + timedelta(hours=2), "13:25")})
    assert r.refresh(conn, EVENING_BEFORE) is True
    after = events_repo.get_event(conn, ev.id)
    assert after.start_time == "11:25"  # the planned time, not the delayed one
    assert after.rev == ev.rev + 1  # saved like an edit


def test_a_typed_time_is_never_replaced(conn):
    ev = add(conn)  # 11:25 typed
    events_repo.update_event(conn, ev.id, {"start_time": "09:30"}, ev.rev)
    refresher({"AC663": flight()})[0].refresh(conn, EVENING_BEFORE)
    assert events_repo.get_event(conn, ev.id).start_time == "09:30"


def test_check_now_fills_the_time_too_and_not_found_leaves_it_empty(conn):
    ev = add_untimed(conn)
    other = add_untimed(conn, title="Room AB 12")
    r, _ = refresher({"AC663": flight()})
    r.check_now(conn, ev, EVENING_BEFORE - timedelta(days=3))
    r.check_now(conn, other, EVENING_BEFORE - timedelta(days=3))
    assert events_repo.get_event(conn, ev.id).start_time == "11:25"
    assert events_repo.get_event(conn, other.id).start_time is None


def test_repeating_events_are_left_alone(conn):
    ev = add_untimed(conn, freq="weekly")
    refresher({"AC663": flight()})[0].refresh(conn, EVENING_BEFORE)
    assert events_repo.get_event(conn, ev.id).start_time is None


def test_a_connection_takes_the_first_flights_time(conn):
    ev = add_untimed(conn, title="AC663 then AC404")
    second = replace(flight(), sched_dep_local="15:10")
    r, _ = refresher({"AC663": flight(), "AC404": second})
    r.check_now(conn, ev, EVENING_BEFORE)
    assert events_repo.get_event(conn, ev.id).start_time == "11:25"


# ---- home airport (owner, 2026-11-20)


def home_flight(revised_arr=None, revised_arr_local=""):
    return FlightInfo(
        "AC664",
        DAY,
        True,
        "Expected",
        DEP,
        None,
        ARR,
        revised_arr,
        sched_dep_local="11:25",
        sched_arr_local="13:09",
        revised_arr_local=revised_arr_local,
        dep_airport="YUL",
        arr_airport="YHZ",
        dep_icao="CYUL",
        arr_icao="CYHZ",
    )


def test_a_flight_landing_at_home_gets_the_arrival_time_and_arrival_status(conn):
    cal.update_calendar(conn, {"home_airport": "yhz"})
    ev = add_untimed(conn, title="AC664")
    late = home_flight(ARR + timedelta(minutes=40), "13:49")
    r, src = refresher({"AC664": late})
    r.refresh(conn, EVENING_BEFORE)
    assert src.calls == [("AC664", DAY)]
    assert events_repo.get_event(conn, ev.id).start_time == "13:09"  # planned arrival
    assert [o.title for o in svc.annotate(conn, _occ(conn))] == ["AC664 (Delayed 1:49p)"]
    shown = svc.stored(conn, "AC664", DAY)
    assert shown.lands_home and shown.word == "Delayed"


def test_without_a_home_airport_the_same_flight_follows_departure(conn):
    ev = add_untimed(conn, title="AC664")
    refresher({"AC664": home_flight()})[0].refresh(conn, EVENING_BEFORE)
    assert events_repo.get_event(conn, ev.id).start_time == "11:25"
    assert svc.stored(conn, "AC664", DAY).lands_home is False


def test_settings_validate_and_store_home_airport_and_limit(conn):
    c = cal.update_calendar(conn, {"home_airport": " cyhz ", "flight_monthly_limit": "150"})
    assert (c.home_airport, c.flight_monthly_limit) == ("CYHZ", 150)
    for bad in ({"home_airport": "Y1Z"}, {"home_airport": "HALIFAX"}):
        with pytest.raises(ValidationFailed):
            cal.update_calendar(conn, bad)
    for bad in ("0", "1001", "lots"):
        with pytest.raises(ValidationFailed):
            cal.update_calendar(conn, {"flight_monthly_limit": bad})
    assert cal.update_calendar(conn, {"home_airport": ""}).home_airport == ""


def test_environment_only_seeds_settings_never_saved(conn):
    assert cal.get_calendar(conn).flight_monthly_limit == 100  # the default
    assert cal.seed_options(conn, {"home_airport": "YHZ", "flight_monthly_limit": 150}) == [
        "home_airport",
        "flight_monthly_limit",
    ]
    cal.update_calendar(conn, {"flight_monthly_limit": "80"})
    assert cal.seed_options(conn, {"home_airport": "YUL", "flight_monthly_limit": 190}) == []
    c = cal.get_calendar(conn)
    assert (c.home_airport, c.flight_monthly_limit) == ("YHZ", 80)
