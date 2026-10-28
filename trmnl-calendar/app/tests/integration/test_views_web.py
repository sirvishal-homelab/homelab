from datetime import UTC, datetime

import pytest
from fastapi.testclient import TestClient

from trmnl_calendar.main import create_app

NOW = datetime(2026, 10, 28, 15, 0, tzinfo=UTC)
API = "/api/v1"


@pytest.fixture
def client(settings):
    with TestClient(create_app(settings, clock=lambda: NOW)) as c:
        yield c


def new_view(client, **kw):
    r = client.post(f"{API}/views", json={"template": "blank", "name": "Work days", **kw})
    assert r.status_code == 201, r.text
    return r.json()


def test_layouts_templates_display(client):
    layouts = {x["id"]: x for x in client.get(f"{API}/layouts").json()}
    assert set(layouts) == {"week_lanes", "twoweek_list"}
    assert "hide_empty_lanes" in [o["key"] for o in layouts["week_lanes"]["options"]]
    assert {t["id"] for t in client.get(f"{API}/templates").json()} >= {"week_lanes", "blank"}
    d = client.get(f"{API}/display").json()
    assert (d["width"], d["height"], d["bit_depth"]) == (800, 480, 2)


def test_view_crud_activate_delete(client):
    views = client.get(f"{API}/views").json()
    assert [v["active"] for v in views] == [True, False]
    v = new_view(client, overrides={"weekdays": ["MO", "TU", "WE", "TH", "FR"]})
    assert v["weekdays"] == ["MO", "TU", "WE", "TH", "FR"] and v["active"] is False
    r = client.patch(f"{API}/views/{v['id']}", json={"name": "Weekdays"})
    assert r.json()["name"] == "Weekdays"
    assert client.post(f"{API}/views/{v['id']}/activate").json()["active"] is True
    gone = client.delete(f"{API}/views/{v['id']}")
    assert gone.status_code == 409 and gone.json()["error"]["code"] == "in_use"
    client.post(f"{API}/views/{views[0]['id']}/activate")
    assert client.delete(f"{API}/views/{v['id']}").status_code == 200
    assert client.get(f"{API}/views/{v['id']}").status_code == 404


def test_view_validation_errors_use_the_error_shape(client):
    r = client.post(f"{API}/views", json={"template": "blank", "name": ""})
    assert r.status_code == 422 and "name" in r.json()["error"]["fields"]
    r = client.patch(f"{API}/views/1", json={"weekdays": []})
    assert r.status_code == 422 and "weekdays" in r.json()["error"]["fields"]


def test_order(client):
    ids = [v["id"] for v in client.get(f"{API}/views").json()]
    r = client.put(f"{API}/views/order", json={"ids": ids[::-1]})
    assert [v["id"] for v in r.json()] == ids[::-1]
    assert client.put(f"{API}/views/order", json={"ids": ids[:1]}).status_code == 422


def test_preview_stores_nothing_and_returns_image_and_fit(client):
    before = client.get(f"{API}/views").json()
    body = {"view": {"layout": "twoweek_list", "weekdays": ["MO", "TU", "WE", "TH", "FR"]}}
    r = client.post(f"{API}/views/preview", json=body)
    assert r.status_code == 200, r.text
    out = r.json()
    assert out["image"]["width"] == 800 and out["image"]["url"].startswith("/render/")
    png = client.get(out["image"]["url"])
    assert png.status_code == 200 and png.content.startswith(b"\x89PNG")
    assert client.get(f"{API}/views").json() == before
    bad = client.post(
        f"{API}/views/preview", json={"view": {"layout": "week_lanes", "weekdays": []}}
    )
    assert bad.status_code == 422


def test_preview_changes_with_settings(client):
    a = client.post(
        f"{API}/views/preview", json={"view": {"layout": "week_lanes", "weekdays": ["MO"]}}
    )
    b = client.post(
        f"{API}/views/preview", json={"view": {"layout": "twoweek_list", "weekdays": ["MO"]}}
    )
    assert a.json()["image"]["url"] != b.json()["image"]["url"]


def test_active_view_drives_current_png(client):
    first = client.get("/render/current.png").content
    client.post(f"{API}/views/2/activate")
    assert client.get("/render/current.png").content != first


def test_views_pages(client):
    assert client.get("/views").status_code == 200
    assert "Two weeks - list" in client.get("/views").text
    page = client.get("/views/new?template=twoweek_list")
    assert page.status_code == 200 and "Everything fits" in page.text
    assert client.get("/views/1/edit").status_code == 200


def test_create_edit_and_activate_through_forms(client):
    r = client.post(
        "/views",
        data={
            "name": "Work week",
            "layout": "week_lanes",
            "weekday": ["MO", "TU", "WE", "TH", "FR"],
            "categories_mode": "all",
            "people_mode": "all",
            "show_times": "on",
            "action": "save_show",
        },
        follow_redirects=False,
    )
    assert r.status_code == 303, r.text
    views = client.get(f"{API}/views").json()
    mine = next(v for v in views if v["name"] == "Work week")
    assert mine["active"] is True and len(mine["weekdays"]) == 5
    r = client.post(
        f"/views/{mine['id']}/edit",
        data={"name": "Work week", "layout": "week_lanes", "weekday": ["MO"], "action": "save"},
        follow_redirects=False,
    )
    assert r.status_code == 303
    assert client.get(f"{API}/views/{mine['id']}").json()["weekdays"] == ["MO"]


def test_form_errors_and_preview_button(client):
    bad = client.post("/views", data={"name": "", "layout": "week_lanes", "weekday": ["MO"]})
    assert bad.status_code == 422 and "Please check" in bad.text
    prev = client.post(
        "/views/1/edit",
        data={
            "name": "Week - lanes",
            "layout": "twoweek_list",
            "weekday": ["MO"],
            "action": "preview",
        },
    )
    assert prev.status_code == 200 and "/render/" in prev.text
    assert client.get(f"{API}/views/1").json()["layout"] == "week_lanes"  # nothing saved


def test_delete_active_view_page_is_refused(client):
    r = client.post("/views/1/delete")
    assert r.status_code == 409


def test_home_names_the_active_view(client):
    assert "Week - lanes" in client.get("/").text
