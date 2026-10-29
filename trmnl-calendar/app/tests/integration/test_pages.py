from datetime import UTC, datetime

import pytest
from fastapi.testclient import TestClient

from trmnl_calendar.main import create_app

NOW = datetime(2026, 10, 28, 1, 0, tzinfo=UTC)


@pytest.fixture
def client(settings):
    with TestClient(create_app(settings, clock=lambda: NOW)) as c:
        yield c


FORM = {"start_date": "2026-10-28", "section_id": "1", "title": "Dentist", "repeat": "none"}


def add(client, **kw):
    r = client.post("/events/new", data={**FORM, **kw}, follow_redirects=False)
    assert r.status_code == 303, r.text
    return r


def test_home_has_preview_and_add_button(client):
    r = client.get("/")
    assert r.status_code == 200 and "/render/current.png" in r.text and "Add event" in r.text


def test_add_event_then_see_it_on_home(client):
    r = add(client, start_time="09:30")
    assert r.headers["location"].startswith("/?saved=")
    page = client.get(r.headers["location"])
    assert "Dentist" in page.text and "9:30a" in page.text and "Added" in page.text


def test_titles_are_escaped(client):
    add(client, title="<b>bold</b>")
    assert "<b>bold</b>" not in client.get("/").text


def test_invalid_form_is_redisplayed_with_messages(client):
    r = client.post("/events/new", data={**FORM, "title": "  "})
    assert r.status_code == 422 and "Enter a short description" in r.text
    assert 'value="2026-10-28"' in r.text


def test_trip_and_repeat_via_the_form(client):
    add(client, title="Cabin", start_date="2026-10-30", end_date="2026-11-01")
    add(client, title="Swim", repeat="biweekly")
    page = client.get("/").text
    assert "Day 2 of 3" in page and "repeats" in page


def test_edit_changes_the_event_and_stale_edit_is_refused(client):
    add(client)
    form = client.get("/events/1/edit")
    assert form.status_code == 200 and 'name="rev" value="1"' in form.text
    ok = client.post(
        "/events/1/edit", data={**FORM, "title": "Dentist 2", "rev": "1"}, follow_redirects=False
    )
    assert ok.status_code == 303
    stale = client.post("/events/1/edit", data={**FORM, "title": "Other", "rev": "1"})
    assert stale.status_code == 409 and "changed this event" in stale.text
    assert "not saved yet" in stale.text
    # their typing is kept, the difference is listed, and the form now carries the new rev
    assert 'value="Other"' in stale.text and "saved now “Dentist 2”" in stale.text
    assert 'name="rev" value="2"' in stale.text
    again = client.post(
        "/events/1/edit", data={**FORM, "title": "Other", "rev": "2"}, follow_redirects=False
    )
    assert again.status_code == 303 and "Other" in client.get("/").text


def test_delete_shows_undo_and_restore_brings_it_back(client):
    add(client)
    r = client.post("/events/1/delete", data={"rev": "1"}, follow_redirects=False)
    assert r.status_code == 303 and r.headers["location"] == "/?deleted=1"
    home = client.get(r.headers["location"]).text
    assert "Undo" in home and "Dentist" in home
    assert "Dentist" in client.get("/deleted").text
    back = client.post("/events/1/restore", follow_redirects=False)
    assert back.status_code == 303
    assert "Coming up" in client.get("/").text and "Nothing in the next" not in client.get("/").text


def test_skip_and_unskip_from_the_edit_page(client):
    add(client, repeat="weekly")
    client.post("/events/1/skip", data={"date": "2026-11-04"}, follow_redirects=False)
    assert "Skipped 2026-11-04" in client.get("/events/1/edit").text
    client.post("/events/1/unskip", data={"date": "2026-11-04"}, follow_redirects=False)
    assert "Skipped 2026-11-04" not in client.get("/events/1/edit").text


def test_unknown_event_is_a_friendly_404(client):
    r = client.get("/events/99/edit")
    assert r.status_code == 404 and "no longer exists" in r.text


def test_current_image_and_hash_route(client):
    add(client)
    r = client.get("/render/current.png")
    assert r.status_code == 200 and r.headers["content-type"] == "image/png"
    assert r.content[:8] == b"\x89PNG\r\n\x1a\n"
    h = r.headers["x-render-hash"]
    again = client.get(f"/render/{h}.png")
    assert again.status_code == 200 and again.content == r.content
    assert "immutable" in again.headers["cache-control"]
    assert client.get("/render/" + "0" * 24 + ".png").status_code == 404
    assert client.get("/render/not-a-hash.png").status_code == 404


def test_static_css_is_served(client):
    assert client.get("/static/app.css").status_code == 200


def test_month_page_shows_events_and_add_links(client):
    add(client, start_date="2026-10-28", title="Dentist")
    page = client.get("/month?y=2026&m=10")
    assert page.status_code == 200
    assert "October 2026" in page.text and "Dentist" in page.text
    assert "/events/new?date=2026-10-28&back=" in page.text


def test_month_page_tolerates_bad_params(client):
    assert client.get("/month?y=abc").status_code in (200, 422)
    assert client.get("/month?y=2026&m=99&day=nope").status_code == 200


def test_form_from_month_returns_to_month(client):
    back = "/month?y=2026&m=10"
    form = client.get("/events/new", params={"date": "2026-10-30", "back": back})
    assert 'name="back"' in form.text and "2026-10-30" in form.text
    r = client.post("/events/new", data={**FORM, "back": back}, follow_redirects=False)
    assert r.status_code == 303
    assert r.headers["location"].startswith("/month?y=2026&m=10&saved=")


def test_back_must_be_same_site(client):
    r = client.post(
        "/events/new", data={**FORM, "back": "https://evil.example/x"}, follow_redirects=False
    )
    assert r.headers["location"].startswith("/?saved=")


def test_form_repeat_ends_after_a_number_of_times(client):
    add(client, repeat="custom", custom_freq="weekly", ends="count", count="3", by_weekday="WE")
    page = client.get("/").text
    assert page.count("Dentist") >= 3 or "3 times" in page
    bad = client.post(
        "/events/new",
        data={**FORM, "repeat": "custom", "custom_freq": "weekly", "ends": "count", "count": "1"},
    )
    assert bad.status_code == 422 and "2 to 999" in bad.text


def test_preset_weekly_with_a_count_really_stops(client):
    add(client, repeat="weekly", ends="count", count="3", start_date="2026-10-31")
    page = client.get("/month?y=2026&m=12").text
    assert "Dentist" not in page  # Oct 31, Nov 7, Nov 14 only
    assert "Dentist" in client.get("/month?y=2026&m=11").text


def test_export_downloads_json_with_events(client):
    add(client, title="Dentist")
    r = client.get("/api/v1/export")
    assert r.status_code == 200 and r.headers["content-type"].startswith("application/json")
    assert 'filename="calendar-export-2026-10-28.json"' in r.headers["content-disposition"]
    assert [e["title"] for e in r.json()["tables"]["events"]] == ["Dentist"]
    assert "/api/v1/export" in client.get("/settings").text


def test_home_screen_icon_and_manifest_are_served(client):
    page = client.get("/").text
    assert "/static/manifest.webmanifest" in page and "apple-touch-icon" in page
    for path in ("/static/manifest.webmanifest", "/static/icon-180.png", "/static/icon-512.png"):
        assert client.get(path).status_code == 200
    assert client.get("/static/icon-192.png").content.startswith(b"\x89PNG")


ROTA = dict(
    repeat="rotating",
    start_date="2026-10-26",
    title="Work",
    start_time="08:00",
    end_time="20:00",
    rot_a=["MO", "TU", "FR", "SA", "SU"],
    rot_b=["WE", "TH"],
)


def test_add_a_rotating_schedule_and_set_days(client):
    add(client, **ROTA)
    month = client.get("/month?y=2026&m=11").text
    assert "Work" in month and "Rotating" not in month
    edit = client.get("/events/1/edit").text
    assert "Rotating weeks: Mon Tue Fri Sat Sun; then Wed Thu" in edit
    back = "/month?y=2026&m=10&day=2026-10-30"
    r = client.post(
        "/events/1/day",
        data={"date": "2026-10-30", "note": "out", "back": back},
        follow_redirects=False,
    )
    assert r.status_code == 303 and r.headers["location"].startswith(back.split("&day")[0])
    page = client.get("/month?y=2026&m=10&day=2026-10-30").text
    assert "Work (out)" in page
    client.post("/events/1/day", data={"date": "2026-10-30", "note": ""})
    assert "Work (out)" not in client.get("/month?y=2026&m=10&day=2026-10-30").text


def test_day_off_and_days_off_range_remove_shifts(client):
    add(client, **ROTA)
    client.post("/events/1/skip", data={"date": "2026-10-31", "back": "/"})
    r = client.post(
        "/events/1/skip",
        data={"date": "2026-11-09", "last": "2026-11-15", "back": "/month"},
        follow_redirects=False,
    )
    assert r.status_code == 303 and "Skipped%207%20days" in r.headers["location"]
    nov = client.get("/month?y=2026&m=11&day=2026-11-10").text
    assert "Nothing on this day" in nov


def test_rotation_and_day_notes_through_the_api(client):
    r = client.post(
        "/api/v1/events",
        json={
            "title": "Work",
            "category_id": 1,
            "start_date": "2026-10-26",
            "repeat": {"rotation": [["MO", "FR"], ["WE"]]},
        },
        headers={"Idempotency-Key": "k1"},
    )
    assert r.status_code == 201, r.text
    assert r.json()["repeat"]["rotation"] == [["MO", "FR"], ["WE"]]
    ok = client.put("/api/v1/events/1/days/2026-10-30", json={"note": "in"})
    assert ok.status_code == 200 and ok.json()["note"] == "in"
    occ = client.get("/api/v1/occurrences?from=2026-10-26&to=2026-11-01").json()
    assert [o["title"] for o in occ if o["date"] == "2026-10-30"] == ["Work (in)"]
    bad = client.put("/api/v1/events/1/days/2026-10-30", json={"note": "maybe"})
    assert bad.status_code == 422
