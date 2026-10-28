from datetime import UTC, datetime

import pytest
from fakes.terminus import FakeTerminus

from trmnl_calendar.config import load_settings
from trmnl_calendar.publish import check
from trmnl_calendar.publish.terminus import TerminusError
from trmnl_calendar.render.pipeline import RenderCache
from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import views as views_repo
from trmnl_calendar.web import render_service

NOW = datetime(2026, 10, 28, 15, 0, tzinfo=UTC)


@pytest.fixture
def world(tmp_path):
    fake = FakeTerminus().start()
    conn = db.connect(tmp_path / "c.db")
    migrate.migrate(conn, tmp_path / "c.db", tmp_path / "b")
    png = render_service.render_current(conn, NOW, RenderCache()).png
    display = views_repo.get_display(conn)

    def settings(**kw):
        env = {
            "DATA_DIR": str(tmp_path),
            "TERMINUS_URL": fake.url,
            "TERMINUS_LOGIN": "cal",
            "TERMINUS_PASSWORD": "pw",
            "PUBLIC_BASE_URL": "http://10.0.0.5:8080",
        }
        env.update(kw)
        return load_settings(env)

    def run(screen_id=None, fetch=None, **kw):
        steps = check.run_check(
            settings(**kw), display, screen_id, fetch=fetch or (lambda u: (200, png))
        )
        return {s.key: s for s in steps}

    yield fake, run
    conn.close()
    fake.stop()


def test_everything_ok_before_first_push(world):
    _, run = world
    s = run()
    assert [v.status for v in s.values()] == ["ok", "ok", "ok", "ok", "ok", "info"]
    assert "800 x 480, 2-bit" in s["model"].detail and "TRMNL" in s["device"].detail
    assert "800 x 480 PNG" in s["image"].detail


def test_screen_in_playlist_alone_is_ok(world):
    fake, run = world
    fake.screens[20] = {"id": 20, "name": "trmnl-calendar"}
    fake.playlists[0]["items"] = [{"screen_id": 20}]
    assert run(screen_id=20)["screen"].status == "ok"


def test_screen_missing_from_playlist_or_sharing_it_warns(world):
    fake, run = world
    fake.screens[20] = {"id": 20, "name": "trmnl-calendar"}
    assert run(screen_id=20)["screen"].status == "warn"
    fake.playlists[0]["items"] = [{"screen_id": 20}, {"screen_id": 8}]
    s = run(screen_id=20)["screen"]
    assert s.status == "warn" and "alternate" in s.detail


def test_deleted_screen_warns(world):
    _, run = world
    assert run(screen_id=99)["screen"].status == "warn"


def test_wrong_password_stops_with_a_fix(world):
    _, run = world
    s = run(TERMINUS_PASSWORD="nope")
    assert s["login"].status == "fail" and "TERMINUS_PASSWORD" in s["login"].fix
    assert s["device"].status == "skip" and s["screen"].status == "skip"


def test_unreachable_terminus(world):
    _, run = world
    s = run(TERMINUS_URL="http://127.0.0.1:9")
    assert s["login"].status == "fail" and "TERMINUS_URL" in s["login"].fix


def test_model_problems(world):
    fake, run = world
    assert run(TERMINUS_MODEL_NAME="nope")["model"].status == "fail"
    assert run(TERMINUS_MODEL_NAME="og_png")["model"].status == "fail"  # 1-bit vs our 2-bit


def test_device_problems(world):
    fake, run = world
    fake.devices[0]["model_id"] = 2
    assert run()["device"].status == "warn"
    fake.devices.clear()
    assert run()["device"].status == "fail"


def test_picture_address_problems_warn_not_fail(world):
    _, run = world

    def refuse(url):
        raise TerminusError("unreachable", "connection refused")

    assert run(fetch=refuse)["image"].status == "warn"
    assert run(fetch=lambda u: (404, b""))["image"].status == "warn"


def test_settings_problems(world):
    _, run = world
    only = run(TERMINUS_PASSWORD="")
    assert list(only) == ["settings"] and "TERMINUS_PASSWORD" in only["settings"].detail
    assert run(PUBLIC_BASE_URL="http://calendar.lan:8080")["settings"].status == "warn"
    assert run(PUBLIC_BASE_URL="calendar")["settings"].status == "fail"
