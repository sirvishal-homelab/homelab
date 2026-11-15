import asyncio
import contextlib
from datetime import UTC, datetime

import pytest

from esp_screen import publish
from esp_screen.storage import db, repo
from tests.fakes.screen import FakeScreen

NOON_UTC = datetime(2026, 11, 15, 12, 0, tzinfo=UTC)  # 08:00 in Halifax (AST, UTC-4)


@pytest.fixture(autouse=True)
def no_waiting(monkeypatch):
    async def instant(_seconds):
        return None

    monkeypatch.setattr(publish.asyncio, "sleep", instant)


def push(settings, fake):
    return asyncio.run(
        publish.push_startup_frame(settings, now=lambda: NOON_UTC, transport=fake.transport)
    )


def test_startup_frame_reaches_screen(migrated):
    fake = FakeScreen()
    assert push(migrated, fake) is True
    assert fake.shown == "/image/status.jpg"
    assert fake.files["/image/status.jpg"][:2] == b"\xff\xd8"
    with contextlib.closing(db.connect(migrated.db_path)) as conn:
        state = repo.get_push_state(conn)
    assert state.last_result == "ok"
    assert state.last_hash


def test_screen_offline_is_recorded_not_raised(migrated):
    fake = FakeScreen(fail_next=10)
    assert push(migrated, fake) is False
    with contextlib.closing(db.connect(migrated.db_path)) as conn:
        state = repo.get_push_state(conn)
    assert state.last_result == "error"
    assert "failed" in state.last_error


def test_no_address_pushes_nothing(migrated):
    with contextlib.closing(db.connect(migrated.db_path)) as conn:
        repo.set_setting(conn, "device_url", "")
    fake = FakeScreen()
    assert push(migrated, fake) is False
    assert fake.calls == []


def test_autoplay_on_is_warned(migrated, caplog):
    push(migrated, FakeScreen(autoplay=1))
    assert "autoplay is ON" in caplog.text


def test_album_answering_ok_does_not_fail_the_push(migrated, caplog):
    """Seen on the real screen 2026-11-15: /album.json answered "OK" right after an upload."""
    caplog.set_level("INFO")
    fake = FakeScreen(album_busy=5)
    assert push(migrated, fake) is True
    with contextlib.closing(db.connect(migrated.db_path)) as conn:
        state = repo.get_push_state(conn)
    assert state.last_result == "ok"
    assert state.last_error is None
    assert "check skipped" in caplog.text


def test_album_read_after_one_busy_answer(migrated, caplog):
    push(migrated, FakeScreen(album_busy=1, autoplay=1))
    assert "autoplay is ON" in caplog.text
