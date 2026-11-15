"""App start-up and /healthz. Needs FastAPI (runs in CI; skipped where it is not installed)."""

import time

import pytest

pytest.importorskip("fastapi")

from fastapi.testclient import TestClient  # noqa: E402

from esp_screen.main import create_app  # noqa: E402
from tests.fakes.screen import FakeScreen  # noqa: E402


def test_startup_migrates_backs_up_and_pushes(settings):
    fake = FakeScreen()
    app = create_app(settings, screen_transport=fake.transport)
    with TestClient(app) as client:
        for _ in range(50):  # the start-up push runs in the background
            r = client.get("/healthz")
            if r.json().get("screen", {}).get("last_result"):
                break
            time.sleep(0.05)
    assert r.status_code == 200
    body = r.json()
    assert body["screen"]["last_result"] == "ok"
    assert fake.shown == "/image/status.jpg"
    assert body["status"] == "ok"
    assert body["schema_version"] == 1
    assert body["screen"]["configured"] is True
    assert body["last_backup_age_hours"] is not None
    assert body["beszel_configured"] is False


def test_rejects_more_than_one_worker(tmp_path):
    from esp_screen.config import load_settings

    s = load_settings({"DATA_DIR": str(tmp_path), "WEB_CONCURRENCY": "2"})
    with pytest.raises(RuntimeError, match="WEB_CONCURRENCY"):
        with TestClient(create_app(s)):
            pass
