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
    assert stale.status_code == 409 and "Someone else changed this event" in stale.text
    assert "Dentist 2" in stale.text


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
