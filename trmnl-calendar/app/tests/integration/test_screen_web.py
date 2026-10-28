import socket
import threading
import time
from datetime import UTC, datetime

import httpx
import pytest
import uvicorn
from fakes.terminus import FakeTerminus
from fastapi.testclient import TestClient

from trmnl_calendar.config import load_settings
from trmnl_calendar.main import create_app
from trmnl_calendar.web import render_service

API = "/api/v1"
NOW = datetime(2026, 10, 28, 15, 0, tzinfo=UTC)


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def wait_for(cond, timeout=15.0):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if cond():
            return True
        time.sleep(0.1)
    return False


@pytest.fixture
def fake():
    f = FakeTerminus().start()
    yield f
    f.stop()


def configured(tmp_path, fake, **extra):
    env = {
        "DATA_DIR": str(tmp_path),
        "TERMINUS_URL": fake.url,
        "TERMINUS_LOGIN": "cal",
        "TERMINUS_PASSWORD": "pw",
        "PUBLIC_BASE_URL": "http://127.0.0.1:9",  # the test client is not reachable by Terminus
        "PUSH_STARTUP_DELAY_SECONDS": "3600",
    }
    env.update(extra)
    return load_settings(env)


# ---- without Terminus


def test_unconfigured_status_push_and_page(settings):
    with TestClient(create_app(settings, clock=lambda: NOW)) as c:
        st = c.get(f"{API}/display/status").json()
        assert st["configured"] is False and st["status"] == "never"
        assert "password" not in str(st).lower()
        r = c.post(f"{API}/display/push")
        assert r.status_code == 409 and r.json()["error"]["code"] == "not_configured"
        page = c.get("/settings").text
        assert "Not set up" in page and "TERMINUS_URL" in page
        check = c.post("/settings/screen/check")
        assert check.status_code == 200 and "Connection check" in check.text
        assert c.post(f"{API}/display/check").json()["ok"] is False


def test_enabled_switch(settings):
    with TestClient(create_app(settings, clock=lambda: NOW)) as c:
        assert c.patch(f"{API}/display", json={"enabled": False}).json()["enabled"] is False
        assert c.get(f"{API}/display/status").json()["enabled"] is False
        assert c.patch(f"{API}/display", json={"enabled": "no"}).status_code == 422
        r = c.post("/settings/screen/enabled", data={"enabled": "1"}, follow_redirects=False)
        assert r.status_code == 303 and c.get(f"{API}/display").json()["enabled"] is True


# ---- with a fake Terminus that cannot reach the test client


def test_check_and_push_failure_surface_terminus_reasons(tmp_path, fake):
    with TestClient(create_app(configured(tmp_path, fake), clock=lambda: NOW)) as c:
        chk = c.post(f"{API}/display/check").json()
        keys = {s["key"]: s["status"] for s in chk["steps"]}
        assert keys["login"] == "ok" and keys["device"] == "ok" and chk["ok"] is True
        r = c.post(f"{API}/display/push")
        assert r.status_code == 502 and r.json()["error"]["code"] == "push_failed"
        st = c.get(f"{API}/display/status").json()
        assert st["status"] == "error" and st["failures"] == 1 and st["last_error"]
        page = c.get("/settings").text
        assert "Could not send the picture" in page
        redirect = c.post("/settings/screen/push", follow_redirects=False)
        assert redirect.status_code == 303 and "screen=failed" in redirect.headers["location"]


def test_status_never_leaks_secrets(tmp_path):
    secret = "s3cret-Pass-9x"
    fake = FakeTerminus(password=secret).start()
    try:
        env = configured(tmp_path, fake, TERMINUS_PASSWORD=secret)
        with TestClient(create_app(env, clock=lambda: NOW)) as c:
            c.post(f"{API}/display/push")  # fails (not reachable), writes an error
            bodies = [
                c.get(f"{API}/display/status").text,
                c.get(f"{API}/display").text,
                c.get("/settings").text,
                c.post("/settings/screen/check").text,
                c.post(f"{API}/display/check").text,
                c.post(f"{API}/display/push").text,
            ]
            assert all(secret not in b for b in bodies)
    finally:
        fake.stop()


# ---- the whole thing, over real sockets


def test_end_to_end_startup_push_then_change_push(tmp_path, fake):
    port = free_port()
    settings = configured(
        tmp_path,
        fake,
        PUBLIC_BASE_URL=f"http://127.0.0.1:{port}",
        PUSH_STARTUP_DELAY_SECONDS="0",
        PUSH_DEBOUNCE_SECONDS="0.2",
    )
    app = create_app(settings)
    server = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning"))
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()
    try:
        assert wait_for(lambda: server.started), "server did not start"
        base = f"http://127.0.0.1:{port}"
        assert wait_for(lambda: bool(fake.stored)), "no startup push reached Terminus"
        sid = next(iter(fake.screens))
        first = fake.stored[sid]
        assert first.startswith(b"\x89PNG")

        today = render_service.today_for("America/Halifax", datetime.now(UTC)).isoformat()
        r = httpx.post(
            f"{base}{API}/events",
            json={"title": "Dentist", "category_id": 1, "start_date": today},
        )
        assert r.status_code == 201
        assert wait_for(lambda: fake.stored[sid] != first), "the edit did not reach Terminus"
        assert httpx.get(f"{base}/render/current.png").content == fake.stored[sid]
        assert fake.calls("POST", "/api/screens") == 1  # patched, never re-created
        st = httpx.get(f"{base}{API}/display/status").json()
        assert st["status"] == "ok" and st["screen_id"] == sid and st["failures"] == 0
    finally:
        server.should_exit = True
        thread.join(timeout=10)


# ---- holidays


def test_holidays_api_and_page(settings):
    with TestClient(create_app(settings, clock=lambda: NOW)) as c:
        regions = c.get(f"{API}/holidays/regions").json()
        assert {"CA", "CA-NS"} == {r["id"] for r in regions}
        listing = c.get(f"{API}/holidays", params={"region": "CA-NS", "years": "2026"}).json()
        heritage = next(i for i in listing["items"] if i["key"] == "heritage_day")
        assert heritage["date"] == "2026-02-16" and heritage["default"] is True
        r = c.post(f"{API}/holidays/import", json={"region": "CA-NS", "years": [2026]})
        assert r.status_code == 201 and r.json()["added"] == 10
        again = c.post(f"{API}/holidays/import", json={"region": "CA-NS", "years": [2026]})
        assert again.status_code == 200 and again.json() == {
            "added": 0,
            "skipped": 10,
            "category": r.json()["category"],
        }
        bad = c.post(f"{API}/holidays/import", json={"region": "XX", "years": [2026]})
        assert bad.status_code == 422 and "region" in bad.json()["error"]["fields"]
        assert c.get(f"{API}/holidays", params={"region": "CA", "years": "1900"}).status_code == 422
        assert c.get("/holidays").status_code == 200
        page = c.get("/holidays", params={"region": "CA-NS", "year": "2026"})
        assert page.status_code == 200 and "Heritage Day" in page.text


def test_holiday_form_import_and_home_banner(settings):
    with TestClient(create_app(settings, clock=lambda: NOW)) as c:
        r = c.post(
            "/holidays/import",
            data={"region": "CA-NS", "h": ["christmas:2026", "boxing_day:2026"]},
            follow_redirects=False,
        )
        assert r.status_code == 303 and "Added%202" in r.headers["location"]
        empty = c.post("/holidays/import", data={"region": "CA-NS"})
        assert empty.status_code == 422
        events = c.get(f"{API}/events", params={"from": "2026-12-01", "to": "2026-12-31"}).json()
        assert {e["title"] for e in events} >= {"Christmas Day", "Boxing Day"}
