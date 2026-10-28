from datetime import UTC, datetime

import pytest
from fakes.imageserver import ImageServer
from fakes.terminus import FakeTerminus

from trmnl_calendar.config import load_settings
from trmnl_calendar.publish.publisher import Publisher
from trmnl_calendar.render.pipeline import RenderCache
from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import events as events_repo
from trmnl_calendar.storage.repositories import push_state as state_repo
from trmnl_calendar.storage.repositories import views as views_repo

NOW = datetime(2026, 10, 28, 15, 0, tzinfo=UTC)


@pytest.fixture
def env(tmp_path):
    fake = FakeTerminus().start()
    cache = RenderCache()
    images = ImageServer(cache)
    settings = load_settings(
        {
            "DATA_DIR": str(tmp_path),
            "TERMINUS_URL": fake.url,
            "TERMINUS_LOGIN": "cal",
            "TERMINUS_PASSWORD": "pw",
            "PUBLIC_BASE_URL": images.url,
        }
    )
    conn = db.connect(settings.db_path)
    migrate.migrate(conn, settings.db_path, settings.backup_dir)
    pub = Publisher(settings, cache, lambda: NOW)
    yield fake, images, pub, conn, settings
    conn.close()
    fake.stop()
    images.stop()


def add(conn, title="Dentist"):
    events_repo.create_event(conn, {"title": title, "section_id": 1, "start_date": "2026-10-28"})


def test_not_configured_does_nothing(tmp_path):
    s = load_settings({"DATA_DIR": str(tmp_path)})
    assert not s.terminus_configured and "pw" not in repr(s)
    assert Publisher(s, RenderCache(), lambda: NOW).run().status == "not_configured"


def test_first_push_creates_the_screen_and_stores_the_exact_image(env):
    fake, images, pub, conn, _ = env
    add(conn)
    out = pub.run()
    assert out.status == "ok" and out.screen_id in fake.screens
    stored = fake.stored[out.screen_id]
    assert stored.startswith(b"\x89PNG") and images.hits == [f"/render/{out.image_hash}.png"]
    st = state_repo.get_state(conn)
    assert (st.last_status, st.last_hash, st.screen_id, st.failures) == (
        "ok",
        out.image_hash,
        out.screen_id,
        0,
    )
    assert st.last_pushed_at == "2026-10-28T15:00:00Z"
    assert fake.screens[out.screen_id]["name"] == "trmnl-calendar"


def test_unchanged_picture_makes_no_terminus_call(env):
    fake, _, pub, conn, _ = env
    pub.run()
    before = len(fake.requests)
    out = pub.run()
    assert out.status == "unchanged" and len(fake.requests) == before
    assert state_repo.get_state(conn).last_status == "unchanged"


def test_change_patches_the_same_screen(env):
    fake, _, pub, conn, _ = env
    first = pub.run()
    add(conn)
    second = pub.run()
    assert second.status == "ok" and second.screen_id == first.screen_id
    assert second.image_hash != first.image_hash
    assert fake.calls("POST", "/api/screens") == 1 and fake.calls("PATCH", "/api/screens/") == 1


def test_force_pushes_even_if_unchanged(env):
    fake, _, pub, _, _ = env
    pub.run()
    assert pub.run(force=True).status == "ok"
    assert fake.calls("PATCH", "/api/screens/") == 1


def test_deleted_screen_is_recreated_once(env):
    fake, _, pub, conn, _ = env
    first = pub.run()
    fake.screens.clear()
    add(conn)
    again = pub.run()
    assert again.status == "ok" and again.screen_id != first.screen_id
    assert state_repo.get_state(conn).screen_id == again.screen_id


def test_existing_screen_with_our_name_is_adopted(env):
    fake, _, pub, _, settings = env
    fake.screens[5] = {"id": 5, "model_id": 12, "name": "trmnl-calendar", "label": "old"}
    out = pub.run()
    assert out.screen_id == 5 and fake.calls("POST", "/api/screens") == 0


def test_failure_is_recorded_and_the_next_run_recovers(env):
    fake, _, pub, conn, _ = env
    pub.run()
    add(conn)
    fake.fail_with = 500
    bad = pub.run()
    st = state_repo.get_state(conn)
    assert bad.status == "error" and st.last_status == "error" and st.failures == 1
    assert "500" in (st.last_error or "")
    good = pub.run()
    st = state_repo.get_state(conn)
    assert good.status == "ok" and st.failures == 0 and st.last_error is None


def test_model_mismatch_refuses_to_push(env):
    fake, _, pub, conn, _ = env
    fake.models[0]["bit_depth"] = 1
    out = pub.run()
    assert out.status == "error" and "2-bit" in out.message and not fake.screens


def test_unknown_model_name(env):
    fake, _, pub, _, _ = env
    fake.models.clear()
    out = pub.run()
    assert out.status == "error" and "TERMINUS_MODEL_NAME" in out.message


def test_wrong_password_reports_auth(env):
    fake, _, pub, conn, settings = env
    bad = Publisher(
        load_settings(
            {
                "DATA_DIR": str(settings.data_dir),
                "TERMINUS_URL": fake.url,
                "TERMINUS_LOGIN": "cal",
                "TERMINUS_PASSWORD": "nope",
                "PUBLIC_BASE_URL": "http://127.0.0.1:1",
            }
        ),
        RenderCache(),
        lambda: NOW,
    )
    out = bad.run()
    assert out.status == "error" and "TERMINUS_PASSWORD" in out.message


def test_disabled_display_does_not_push(env):
    fake, _, pub, conn, _ = env
    conn.execute("UPDATE displays SET enabled = 0")
    conn.commit()
    assert pub.run().status == "disabled" and not fake.requests


def test_active_view_is_what_gets_pushed(env):
    fake, _, pub, conn, _ = env
    a = pub.run()
    views_repo.activate_view(conn, 2)
    b = pub.run()
    assert b.status == "ok" and b.image_hash != a.image_hash
