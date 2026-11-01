from datetime import UTC, datetime

import pytest
from fastapi.testclient import TestClient

from trmnl_calendar.main import create_app

NOW = datetime(2026, 10, 28, 1, 0, tzinfo=UTC)
API = "/api/v1"


@pytest.fixture
def client(settings):
    with TestClient(create_app(settings, clock=lambda: NOW)) as c:
        yield c


def person(client, name="Vishal", label="V"):
    r = client.post(f"{API}/people", json={"name": name, "label": label})
    assert r.status_code == 201, r.text
    return r.json()


def event(client, **kw):
    body = {"title": "Dentist", "category_id": 1, "start_date": "2026-10-28", **kw}
    r = client.post(f"{API}/events", json=body)
    assert r.status_code == 201, r.text
    return r.json()


# ---- pages


def test_manage_pages_load(client):
    for path in ("/categories", "/people", "/settings", "/who"):
        assert client.get(path).status_code == 200, path


def test_add_category_page_then_listed(client):
    r = client.post("/categories", data={"name": "Outings", "placement": "lane"})
    assert r.status_code == 200 or r.history  # followed redirect
    assert "Outings" in client.get("/categories").text


def test_duplicate_category_is_a_form_error(client):
    client.post("/categories", data={"name": "Outings"})
    r = client.post("/categories", data={"name": "outings"})
    assert r.status_code == 422


def test_add_person_page(client):
    client.post("/people", data={"name": "Vishal", "label": "V"})
    assert "Vishal" in client.get("/people").text
    r = client.post("/people", data={"name": "Other", "label": "v"})
    assert r.status_code == 422


def test_settings_save_and_reject_bad_timezone(client):
    r = client.post(
        "/settings",
        data={
            "name": "Home",
            "timezone": "America/Toronto",
            "week_start": "mon",
            "time_format": "24",
        },
    )
    assert r.status_code == 200
    assert client.get(f"{API}/calendar").json()["timezone"] == "America/Toronto"
    bad = client.post(
        "/settings",
        data={"name": "Home", "timezone": "Nowhere/Land", "week_start": "mon", "time_format": "24"},
    )
    assert bad.status_code == 422


def test_who_cookie_and_open_redirect_guard(client):
    p = person(client)
    r = client.post(
        "/who",
        data={"person_id": str(p["id"]), "next": "https://evil.example/"},
        follow_redirects=False,
    )
    assert r.status_code == 303 and r.headers["location"] == "/"
    assert client.get(f"{API}/me").json()["person"]["label"] == "V"


def test_who_prompt_shown_until_chosen(client):
    person(client)
    assert "/who" in client.get("/").text


def test_person_filter_on_home(client):
    a, b = person(client), person(client, "Wife", "W")
    event(client, title="Mine", people={"ids": [a["id"]]})
    event(client, title="Hers", people={"ids": [b["id"]]})
    event(client, title="Everyone")
    page = client.get(f"/?person={a['id']}").text
    assert "Mine" in page and "Everyone" in page and "Hers" not in page


def test_delete_category_with_events_requires_move(client):
    c = client.post(f"{API}/categories", json={"name": "Outings"}).json()
    event(client, category_id=c["id"])
    r = client.post(f"/categories/{c['id']}/delete", data={})
    assert r.status_code == 409
    r = client.post(f"/categories/{c['id']}/delete", data={"move_to": "1"})
    assert all(x["id"] != c["id"] for x in client.get(f"{API}/categories").json())


# ---- API


def test_category_api_crud_and_order(client):
    c = client.post(f"{API}/categories", json={"name": "Outings"}).json()
    assert client.patch(f"{API}/categories/{c['id']}", json={"weight": 2}).json()["weight"] == 2
    ids = [x["id"] for x in client.get(f"{API}/categories").json()]
    r = client.put(f"{API}/categories/order", json={"ids": ids[::-1]})
    assert [x["id"] for x in r.json()] == ids[::-1]
    assert client.put(f"{API}/categories/order", json={"ids": ids[:1]}).status_code == 422


def test_last_category_cannot_be_deleted(client):
    only = client.get(f"{API}/categories").json()
    if len(only) == 1:
        r = client.delete(f"{API}/categories/{only[0]['id']}")
        assert r.status_code == 409 and r.json()["error"]["code"] == "in_use"


def test_people_api_and_in_use(client):
    p = person(client)
    assert client.patch(f"{API}/people/{p['id']}", json={"name": "Vish"}).json()["name"] == "Vish"
    event(client, people={"ids": [p["id"]]})
    r = client.delete(f"{API}/people/{p['id']}")
    assert r.status_code == 409 and r.json()["error"]["code"] == "in_use"
    assert client.patch(f"{API}/people/{p['id']}", json={"active": False}).json()["active"] is False


def test_event_people_owner_and_updated_by(client):
    a, b = person(client), person(client, "Wife", "W")
    client.put(f"{API}/me", json={"person_id": a["id"]})
    ev = event(client, people={"ids": [a["id"], b["id"]], "owner_id": b["id"]})
    assert ev["people"]["ids"][0] == b["id"] and ev["people"]["owner_id"] == b["id"]
    assert ev["updated_by"]["label"] == "V"


def test_calendar_api_patch(client):
    r = client.patch(f"{API}/calendar", json={"week_start": "mon", "time_format": "24"})
    assert r.status_code == 200 and r.json()["week_start"] == "mon"
    assert client.patch(f"{API}/calendar", json={"timezone": "Bad/Zone"}).status_code == 422


def test_badges_change_the_rendered_image(client):
    before = client.get("/render/current.png").content
    p = person(client)
    event(client, people={"ids": [p["id"]], "owner_id": p["id"]})
    assert client.get("/render/current.png").content != before


# ---- look (auto / light / dark), remembered per device


def test_look_defaults_to_auto_and_sets_no_theme_attribute(client):
    html = client.get("/settings").text
    assert "data-theme" not in html
    assert 'aria-pressed="true">Auto' in html


@pytest.mark.parametrize("choice", ["light", "dark"])
def test_look_choice_is_remembered_in_a_cookie(client, choice):
    r = client.post("/settings/theme", data={"theme": choice}, follow_redirects=False)
    assert r.status_code == 303
    assert r.headers["location"] == "/settings#look"
    assert f"theme={choice}" in r.headers["set-cookie"]
    html = client.get("/settings").text
    assert f'<html lang="en" data-theme="{choice}">' in html
    assert f'aria-pressed="true">{choice.capitalize()}' in html


def test_look_choice_applies_to_every_page(client):
    client.post("/settings/theme", data={"theme": "dark"})
    for path in ("/", "/categories", "/people", "/who"):
        assert 'data-theme="dark"' in client.get(path).text, path


def test_look_ignores_unknown_values(client):
    client.post("/settings/theme", data={"theme": "dark"})
    client.post("/settings/theme", data={"theme": "neon"})  # not stored; the old choice stays
    assert 'data-theme="dark"' in client.get("/settings").text
    client.cookies.set("theme", "neon")
    assert "data-theme" not in client.get("/settings").text


def test_look_does_not_trigger_a_push(client):
    from trmnl_calendar.main import NO_PUSH_PATHS

    assert "/settings/theme" in NO_PUSH_PATHS


def test_css_and_js_urls_carry_a_version_so_browsers_do_not_keep_old_files(client):
    import re

    html = client.get("/settings").text
    assert re.search(r'/static/app\.css\?v=[0-9a-f]{10}"', html)
    assert re.search(r'/static/app\.js\?v=[0-9a-f]{10}"', html)
    assert client.get("/static/app.css?v=anything").status_code == 200


# ---- restyle: bottom bar on phones, More page, bundled font


def test_more_page_lists_everything_without_a_tab(client):
    r = client.get("/more")
    assert r.status_code == 200
    for target in ("/categories", "/people", "/settings", "/who", "/deleted", "/api/v1/export"):
        assert f'href="{target}' in r.text, target


def test_nav_has_phone_and_wide_links_and_marks_the_current_page(client):
    html = client.get("/more").text
    assert 'class="only-narrow" href="/more" aria-current=page' in html
    assert 'class="only-wide" href="/settings"' in html
    # Categories, People and Settings belong under More on a phone
    assert 'href="/more" aria-current=page' in client.get("/settings").text
    assert 'href="/month" aria-current=page' in client.get("/month").text


def test_font_is_served_by_the_app_not_a_third_party(client):
    css = client.get("/static/app.css").text
    assert "@font-face" in css and "/static/fonts/DMSans-Regular.woff" in css
    assert "fonts.googleapis.com" not in css
    for name in ("Regular", "Medium", "Bold"):
        r = client.get(f"/static/fonts/DMSans-{name}.woff")
        assert r.status_code == 200 and len(r.content) > 5000, name
    assert client.get("/static/fonts/OFL.txt").status_code == 200
