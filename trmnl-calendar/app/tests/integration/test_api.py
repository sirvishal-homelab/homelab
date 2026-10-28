from datetime import UTC, datetime

import pytest
from fastapi.testclient import TestClient

from trmnl_calendar.main import create_app

NOW = datetime(2026, 10, 28, 1, 0, tzinfo=UTC)  # the evening of 27 Oct in Halifax
BASE = "/api/v1"


@pytest.fixture
def client(settings):
    with TestClient(create_app(settings, clock=lambda: NOW)) as c:
        yield c


def post(client, **kw):
    body = {"title": "Dentist", "category_id": 1, "start_date": "2026-10-28"}
    body.update(kw)
    r = client.post(f"{BASE}/events", json=body)
    assert r.status_code == 201, r.text
    return r.json()


def test_create_get_and_repeat_text(client):
    ev = post(client, repeat={"freq": "weekly", "interval": 2})
    assert ev["rev"] == 1 and ev["repeat_text"] == "Repeats every 2 weeks on Wednesday"
    assert client.get(f"{BASE}/events/{ev['id']}").json()["title"] == "Dentist"


def test_patch_needs_if_match_and_detects_conflicts(client):
    ev = post(client)
    url = f"{BASE}/events/{ev['id']}"
    r = client.patch(url, json={"title": "A"})
    assert r.status_code == 428 and r.json()["error"]["code"] == "precondition_required"
    r = client.patch(url, json={"title": "A"}, headers={"If-Match": "1"})
    assert r.status_code == 200 and r.json()["rev"] == 2 and r.json()["start_date"] == "2026-10-28"
    r = client.patch(url, json={"title": "B"}, headers={"If-Match": '"1"'})
    assert r.status_code == 409 and r.json()["error"]["code"] == "edit_conflict"
    assert r.json()["error"]["current_rev"] == 2


def test_validation_errors_use_the_single_shape(client):
    r = client.post(f"{BASE}/events", json={})
    err = r.json()["error"]
    assert r.status_code == 422 and err["code"] == "validation_failed"
    assert err["fields"]["title"] == "Enter a short description"


def test_malformed_json_is_a_400_in_the_same_shape(client):
    r = client.post(f"{BASE}/events", content="{bad", headers={"Content-Type": "application/json"})
    assert r.status_code == 400 and r.json()["error"]["code"] == "bad_request"


def test_idempotency_key_prevents_duplicates(client):
    body = {"title": "Once", "category_id": 1, "start_date": "2026-10-28"}
    h = {"Idempotency-Key": "abc-123"}
    a = client.post(f"{BASE}/events", json=body, headers=h)
    b = client.post(f"{BASE}/events", json=body, headers=h)
    assert a.status_code == b.status_code == 201 and a.json()["id"] == b.json()["id"]
    listed = client.get(f"{BASE}/events", params={"from": "2026-10-01", "to": "2026-10-31"})
    assert len(listed.json()) == 1


def test_delete_and_restore(client):
    ev = post(client)
    url = f"{BASE}/events/{ev['id']}"
    assert client.delete(url).status_code == 428
    gone = client.delete(url, headers={"If-Match": "1"})
    assert gone.status_code == 200 and gone.json()["deleted_at"]
    r = client.get(url)
    assert r.status_code == 404 and r.json()["error"]["code"] == "not_found"
    assert [e["id"] for e in client.get(f"{BASE}/events", params={"deleted": "true"}).json()] == [
        ev["id"]
    ]
    back = client.post(f"{url}/restore")
    assert back.status_code == 200 and back.json()["deleted_at"] is None


def test_skip_and_occurrences(client):
    ev = post(client, repeat={"freq": "weekly"})
    q = {"from": "2026-10-25", "to": "2026-11-14"}
    dates = [o["date"] for o in client.get(f"{BASE}/occurrences", params=q).json()]
    assert dates == ["2026-10-28", "2026-11-04", "2026-11-11"]
    r = client.post(f"{BASE}/events/{ev['id']}/skips", json={"date": "2026-11-04"})
    assert r.status_code == 201 and r.json()["skipped"] == ["2026-11-04"]
    dates = [o["date"] for o in client.get(f"{BASE}/occurrences", params=q).json()]
    assert dates == ["2026-10-28", "2026-11-11"]
    client.delete(f"{BASE}/events/{ev['id']}/skips/2026-11-04")
    assert len(client.get(f"{BASE}/occurrences", params=q).json()) == 3


def test_trip_occurrences_carry_day_numbers(client):
    post(client, start_date="2026-10-30", end_date="2026-11-01", title="Cabin")
    occ = client.get(
        f"{BASE}/occurrences", params={"from": "2026-10-30", "to": "2026-11-02"}
    ).json()
    assert [(o["date"], o["day_number"], o["day_count"]) for o in occ] == [
        ("2026-10-30", 1, 3),
        ("2026-10-31", 2, 3),
        ("2026-11-01", 3, 3),
    ]


def test_window_parameters_are_required(client):
    r = client.get(f"{BASE}/events")
    assert r.status_code == 400 and r.json()["error"]["code"] == "bad_request"


def test_categories_and_unknown_route(client):
    cats = client.get(f"{BASE}/categories").json()
    assert [c["name"] for c in cats] == ["Events", "Work", "Lunch", "Dinner"]
    r = client.get(f"{BASE}/nope")
    assert r.status_code == 404 and r.json()["error"]["code"] == "not_found"


def test_repeat_count_through_the_api(client):
    ev = post(client, repeat={"freq": "weekly", "count": 6})
    assert ev["repeat"]["count"] == 6 and ev["repeat_text"].endswith(", 6 times")
    r = client.patch(
        f"{BASE}/events/{ev['id']}",
        json={"repeat": {"until": "2026-12-01"}},
        headers={"If-Match": '"1"'},
    )
    assert r.status_code == 200 and r.json()["repeat"]["count"] is None
    bad = client.post(
        f"{BASE}/events",
        json={
            "title": "x",
            "category_id": 1,
            "start_date": "2026-10-28",
            "repeat": {"freq": "weekly", "count": 1},
        },
    )  # noqa: E501
    assert bad.status_code == 422 and "count" in bad.json()["error"]["fields"]
