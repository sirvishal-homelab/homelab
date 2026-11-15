"""App start-up and /healthz. Needs FastAPI (runs in CI; skipped where it is not installed)."""

import time

import pytest

pytest.importorskip("fastapi")

from fastapi.testclient import TestClient  # noqa: E402

from esp_screen.main import create_app  # noqa: E402
from tests.fakes.screen import FakeScreen  # noqa: E402


def test_startup_migrates_backs_up_and_pushes(tmp_path):
    from esp_screen.config import load_settings

    fake = FakeScreen()
    with fake.serve() as url:
        settings = load_settings({"DATA_DIR": str(tmp_path), "SCREEN_URL": url})
        body = _start_and_wait(settings)
    assert body["status"] == "ok"
    assert body["schema_version"] == 1
    assert body["screen"]["configured"] is True
    assert body["screen"]["last_result"] == "ok"
    assert body["last_backup_age_hours"] is not None
    assert body["beszel_configured"] is False
    assert fake.shown == "/image/status.jpg"


def _start_and_wait(settings) -> dict:
    with TestClient(create_app(settings)) as client:
        for _ in range(50):  # the start-up push runs in the background
            r = client.get("/healthz")
            if r.json().get("screen", {}).get("last_result"):
                break
            time.sleep(0.05)
    assert r.status_code == 200
    return r.json()


def test_rejects_more_than_one_worker(tmp_path):
    from esp_screen.config import load_settings

    s = load_settings({"DATA_DIR": str(tmp_path), "WEB_CONCURRENCY": "2"})
    with pytest.raises(RuntimeError, match="WEB_CONCURRENCY"):
        with TestClient(create_app(s)):
            pass


def test_live_overview_reported_in_healthz(tmp_path):
    from esp_screen.config import load_settings
    from tests.fakes.beszel import PASSWORD, USER, FakeHub

    hub, fake = FakeHub(), FakeScreen()
    with fake.serve() as url:
        settings = load_settings(
            {
                "DATA_DIR": str(tmp_path),
                "SCREEN_URL": url,
                "BESZEL_URL": "http://beszel:8090",
                "BESZEL_USER": USER,
                "BESZEL_PASSWORD": PASSWORD,
                "FOCUS_SYSTEMS": "node-a,node-b,nas",
            }
        )
        with TestClient(create_app(settings, beszel_transport=hub.transport)) as client:
            for _ in range(100):
                body = client.get("/healthz").json()
                if "overview" in body:
                    break
                time.sleep(0.05)
    assert body["beszel"]["last_error"] is None
    assert body["overview"] == {
        "focus": ["node-a:up", "node-b:up", "nas:up"],
        "containers": "9/9",
        "other_systems": "3/3",
        "paused": 1,
    }
