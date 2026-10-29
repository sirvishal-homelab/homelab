from datetime import UTC, datetime, timedelta

import pytest

from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import device_samples as ds
from trmnl_calendar.storage.repositories import push_state as state_repo

NOW = datetime(2026, 10, 29, 12, 0, tzinfo=UTC)


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def dev(volts=3.9, synced="2026-10-29T11:50:00+00:00", **kw):
    return {
        "id": 7,
        "model_id": 12,
        "battery_voltage": volts,
        "wifi_signal": -40,
        "synced_at": synced,
        "refresh_rate": 600,
        **kw,
    }


def test_a_reading_is_stored_in_utc_text(conn):
    assert ds.record(conn, [dev(synced="2026-10-29T08:50:00-03:00")], 12, NOW)
    [row] = ds.since(conn, NOW - timedelta(days=1))
    assert row["voltage"] == 3.9 and row["wifi"] == -40
    assert row["seen_at"] == "2026-10-29T11:50:00Z" and row["at"] == "2026-10-29T12:00:00Z"


def test_missing_or_silly_voltage_is_ignored(conn):
    for bad in (None, "3.9", True, 9.9, 0.5):
        assert not ds.record(conn, [dev(volts=bad)], 12, NOW)
    assert not ds.record(conn, [], 12, NOW)
    assert ds.since(conn, NOW - timedelta(days=1)) == []


def test_same_check_in_or_too_soon_is_not_stored_twice(conn):
    assert ds.record(conn, [dev()], 12, NOW)
    assert not ds.record(conn, [dev()], 12, NOW + timedelta(minutes=30))  # no new check-in
    assert not ds.record(
        conn, [dev(synced="2026-10-29T12:01:00+00:00")], 12, NOW + timedelta(minutes=2)
    )
    assert ds.record(
        conn, [dev(synced="2026-10-29T12:11:00+00:00")], 12, NOW + timedelta(minutes=15)
    )
    assert len(ds.since(conn, NOW - timedelta(days=1))) == 2


def test_the_right_device_is_used_and_old_rows_are_pruned(conn):
    other = dev(volts=3.3, model_id=99)
    ds.record(conn, [other, dev(volts=4.0)], 12, NOW - timedelta(days=100))
    ds.record(conn, [other, dev(volts=3.8, synced="2026-10-29T11:59:00+00:00")], 12, NOW)
    rows = ds.since(conn, NOW - timedelta(days=200))
    assert [r["voltage"] for r in rows] == [3.8]  # the 100-day-old row was removed


def test_remember_device_keeps_refresh_and_takes_a_reading(conn):
    state_repo.remember_device(conn, [dev()], 12)
    assert state_repo.device_refresh_seconds(conn) == 600
    assert len(ds.since(conn, datetime.now(UTC) - timedelta(days=1))) == 1
