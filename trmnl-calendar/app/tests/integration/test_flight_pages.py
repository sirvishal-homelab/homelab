"""The event page's flight panel, Check now, and the credit line (docs/21 section 4)."""

from contextlib import closing
from datetime import UTC, date, datetime

import pytest
from fastapi.testclient import TestClient

from trmnl_calendar.config import load_settings
from trmnl_calendar.domain.flights import FlightInfo
from trmnl_calendar.main import create_app
from trmnl_calendar.storage import db
from trmnl_calendar.storage.repositories import events as events_repo

NOW = datetime(2026, 11, 15, 23, 0, tzinfo=UTC)  # the evening before the flight, in Halifax
DAY = date(2026, 11, 16)
DEP = datetime(2026, 11, 16, 15, 25, tzinfo=UTC)


class FakeSource:
    def __init__(self):
        self.calls = []

    def lookup(self, code, day, home=""):
        self.calls.append((code, day))
        return FlightInfo(
            code, day, True, "Delayed", DEP, None, None, None, "11:25", "", "YHZ", "YUL", "23"
        )


def make_client(tmp_path, on):
    env = {"DATA_DIR": str(tmp_path), "PUSH_STARTUP_DELAY_SECONDS": "3600"}
    if on:
        env.update({"FLIGHT_STATUS": "on", "FLIGHT_API_KEY": "test-key"})
    app = create_app(load_settings(env), clock=lambda: NOW)
    if on:
        app.state.flights.source = FakeSource()
    return app


def add_flight(settings_path, title="AC663"):
    with closing(db.connect(settings_path)) as conn:
        return events_repo.create_event(
            conn,
            {"title": title, "section_id": 1, "start_date": DAY.isoformat(), "start_time": "11:25"},
        )


@pytest.fixture
def on(tmp_path):
    app = make_client(tmp_path, on=True)
    with TestClient(app) as c:
        yield c, app


@pytest.fixture
def off(tmp_path):
    app = make_client(tmp_path, on=False)
    with TestClient(app) as c:
        yield c, app


def test_key_is_not_shown_in_settings_repr(tmp_path):
    s = load_settings({"DATA_DIR": str(tmp_path), "FLIGHT_STATUS": "on", "FLIGHT_API_KEY": "k123"})
    assert s.flight_enabled and "k123" not in repr(s)


def test_off_by_default_and_the_page_says_how_to_turn_it_on(off):
    client, app = off
    assert app.state.flights is None
    ev = add_flight(app.state.settings.db_path)
    page = client.get(f"/events/{ev.id}/edit").text
    assert "Flight status" in page and "FLIGHT_STATUS=on" in page and "Check now" not in page


def test_events_without_a_code_have_no_panel(on):
    client, app = on
    with closing(db.connect(app.state.settings.db_path)) as conn:
        ev = events_repo.create_event(
            conn, {"title": "Dentist", "section_id": 1, "start_date": DAY.isoformat()}
        )
    assert "Flight status" not in client.get(f"/events/{ev.id}/edit").text


def test_check_now_looks_up_and_shows_the_detail_and_credit(on):
    client, app = on
    ev = add_flight(app.state.settings.db_path)
    page = client.get(f"/events/{ev.id}/edit").text
    assert "Check now" in page and "not looked up yet" in page
    r = client.post(f"/events/{ev.id}/flights/check", follow_redirects=False)
    assert r.status_code == 303 and r.headers["location"].endswith("/edit#flights")
    assert app.state.flights.source.calls == [("AC663", DAY)]
    page = client.get(f"/events/{ev.id}/edit").text
    assert "Delayed" in page and "YHZ" in page and "gate 23" in page
    assert "Flight data: AeroDataBox" in page


def test_home_and_api_show_the_word_and_the_credit(on):
    client, app = on
    ev = add_flight(app.state.settings.db_path)
    client.post(f"/events/{ev.id}/flights/check")
    home = client.get("/").text
    assert "AC663 (Delayed)" in home and "Flight data: AeroDataBox" in home
    rows = client.get("/api/v1/occurrences?from=2026-11-16&to=2026-11-16").json()
    assert [(r["title"], r["flight"]) for r in rows] == [("AC663 (Delayed)", "Delayed")]


def test_check_now_on_an_unknown_event_is_404(on):
    client, _ = on
    assert client.post("/events/999/flights/check").status_code == 404


def test_settings_page_saves_home_airport_and_lookup_limit(on):
    client, _ = on
    page = client.get("/settings").text
    assert 'name="home_airport"' in page and 'name="flight_monthly_limit"' in page
    form = {
        "name": "Home",
        "timezone": "America/Halifax",
        "week_start": "sun",
        "time_format": "12",
        "home_airport": "yhz",
        "flight_monthly_limit": "150",
    }
    assert client.post("/settings", data=form).status_code == 200
    got = client.get("/api/v1/calendar").json()
    assert (got["home_airport"], got["flight_monthly_limit"]) == ("YHZ", 150)
    bad = client.post("/settings", data={**form, "home_airport": "HALIFAX"})
    assert bad.status_code == 422 and "airport code" in bad.text


def test_environment_seeds_the_settings_once(tmp_path):
    env = {
        "DATA_DIR": str(tmp_path),
        "PUSH_STARTUP_DELAY_SECONDS": "3600",
        "FLIGHT_HOME_AIRPORT": "yhz",
        "FLIGHT_MONTHLY_LIMIT": "150",
    }
    with TestClient(create_app(load_settings(env), clock=lambda: NOW)) as c:
        got = c.get("/api/v1/calendar").json()
        assert (got["home_airport"], got["flight_monthly_limit"]) == ("YHZ", 150)
        c.patch("/api/v1/calendar", json={"flight_monthly_limit": 90})
    with TestClient(create_app(load_settings(env), clock=lambda: NOW)) as c:
        assert c.get("/api/v1/calendar").json()["flight_monthly_limit"] == 90  # saved value wins


def test_event_page_shows_arrival_for_a_flight_landing_home(on):
    client, app = on
    client.patch("/api/v1/calendar", json={"home_airport": "YUL"})  # FakeSource lands at YUL
    ev = add_flight(app.state.settings.db_path)
    client.post(f"/events/{ev.id}/flights/check")
    page = client.get(f"/events/{ev.id}/edit").text
    assert "lands at home" in page
