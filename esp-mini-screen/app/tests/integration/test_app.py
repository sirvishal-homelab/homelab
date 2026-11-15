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


# --- S-4: settings page ------------------------------------------------------------------


def _live_settings(tmp_path, url, focus="node-a,node-b,nas"):
    from esp_screen.config import load_settings
    from tests.fakes.beszel import PASSWORD, USER

    return load_settings(
        {
            "DATA_DIR": str(tmp_path),
            "SCREEN_URL": url,
            "BESZEL_URL": "http://beszel:8090",
            "BESZEL_USER": USER,
            "BESZEL_PASSWORD": PASSWORD,
            "FOCUS_SYSTEMS": focus,
        }
    )


def _wait_for(client, key):
    for _ in range(100):
        body = client.get("/healthz").json()
        if key in body:
            return body
        time.sleep(0.05)
    raise AssertionError(f"{key} never appeared in /healthz")


def test_settings_page_save_changes_screen(tmp_path):
    from esp_screen.storage import db, repo
    from tests.fakes.beszel import FakeHub

    hub, fake = FakeHub(), FakeScreen()
    with fake.serve() as url:
        s = _live_settings(tmp_path, url)
        with TestClient(create_app(s, beszel_transport=hub.transport)) as client:
            _wait_for(client, "overview")
            page = client.get("/")
            assert page.status_code == 200
            assert "Focus systems" in page.text and "node-a" in page.text
            import contextlib

            with contextlib.closing(db.connect(s.db_path)) as c:
                ids = {p.beszel_name: p.beszel_id for p in repo.list_systems(c)}
            form = {
                "warn_pct": "70",
                "crit_pct": "85",
                "poll_s": "60",
                "alert_flip_s": "10",
                "unreachable_grace_min": "5",
                "container_missing_polls": "2",
                "device_url": url,
                f"pos_{ids['nas']}": "1",
                f"name_{ids['nas']}": "NAS",
                f"pos_{ids['dns']}": "2",
            }
            r = client.post("/settings", data=form, follow_redirects=False)
            assert r.status_code == 303 and r.headers["location"] == "/?saved=1"
            for _ in range(100):
                body = client.get("/healthz").json()
                if body["overview"]["focus"] == ["NAS:up", "dns:up"]:
                    break
                time.sleep(0.05)
            assert body["overview"]["focus"] == ["NAS:up", "dns:up"]
            assert "Saved." in client.get("/?saved=1").text
            with contextlib.closing(db.connect(s.db_path)) as c:
                assert repo.get_setting(c, "warn_pct") == 70


def test_settings_errors_are_shown_and_nothing_saved(tmp_path):
    import contextlib

    from esp_screen.storage import db, repo
    from tests.fakes.beszel import FakeHub

    hub, fake = FakeHub(), FakeScreen()
    with fake.serve() as url:
        s = _live_settings(tmp_path, url)
        with TestClient(create_app(s, beszel_transport=hub.transport)) as client:
            _wait_for(client, "overview")
            r = client.post("/settings", data={"poll_s": "1", "warn_pct": "95", "crit_pct": "90"})
            assert r.status_code == 400
            assert "Not saved" in r.text and "between 15 and 600" in r.text
            with contextlib.closing(db.connect(s.db_path)) as c:
                assert repo.get_setting(c, "poll_s") == 60


def test_preview_push_now_and_forget(tmp_path):
    import contextlib

    from esp_screen.storage import db, repo
    from tests.fakes.beszel import FakeHub

    hub, fake = FakeHub(), FakeScreen()
    with fake.serve() as url:
        s = _live_settings(tmp_path, url)
        with TestClient(create_app(s, beszel_transport=hub.transport)) as client:
            _wait_for(client, "overview")
            img = client.get("/preview.jpg")
            assert img.status_code == 200 and img.headers["content-type"] == "image/jpeg"
            uploads = sum(c.startswith("POST") for c in fake.calls)
            assert client.post("/push", follow_redirects=False).status_code == 303
            for _ in range(100):
                if sum(c.startswith("POST") for c in fake.calls) > uploads:
                    break
                time.sleep(0.05)
            assert sum(c.startswith("POST") for c in fake.calls) > uploads
            with contextlib.closing(db.connect(s.db_path)) as c:
                sid = repo.list_containers(c)[0].system_id
                repo.save_container_memory(c, {(sid, "removed-for-good"): 5})
            r = client.post(
                "/containers/forget",
                data={"key": f"{sid}:removed-for-good"},
                follow_redirects=False,
            )
            assert r.status_code == 303 and r.headers["location"] == "/?forgot=1"
            with contextlib.closing(db.connect(s.db_path)) as c:
                names = {k.name for k in repo.list_containers(c)}
            assert "removed-for-good" not in names
            assert names  # the running ones are untouched
