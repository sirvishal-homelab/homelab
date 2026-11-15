import contextlib
import sqlite3

import pytest

from esp_screen.storage import backup, db, migrate, repo


def test_migrations_are_contiguous():
    versions = [v for v, _, _ in migrate.available()]
    assert versions == list(range(1, len(versions) + 1))


def test_fresh_database_migrates_without_backup(settings):
    with contextlib.closing(db.connect(settings.db_path)) as conn:
        applied = migrate.migrate(conn, settings.db_path, settings.backup_dir)
        assert applied == ["0001_baseline.sql"]
        assert migrate.current_version(conn) == len(migrate.available())
        assert migrate.migrate(conn, settings.db_path, settings.backup_dir) == []
    assert not settings.backup_dir.exists()  # nothing to protect on a new database


def test_default_settings(migrated):
    with contextlib.closing(db.connect(migrated.db_path)) as conn:
        assert repo.get_setting(conn, "poll_s") == 60
        assert repo.get_setting(conn, "alert_flip_s") == 10
        assert repo.get_setting(conn, "warn_pct") == 75
        assert repo.get_setting(conn, "crit_pct") == 90
        assert repo.get_setting(conn, "unreachable_grace_min") == 5
        assert repo.get_setting(conn, "container_missing_polls") == 2
        assert repo.get_setting(conn, "device_url") == "http://screen.test"


def test_seed_does_not_overwrite_owner_choice(migrated):
    with contextlib.closing(db.connect(migrated.db_path)) as conn:
        repo.set_setting(conn, "device_url", "http://chosen.test")
        assert repo.seed_device_url(conn, "http://env.test") == "http://chosen.test"
        assert repo.get_setting(conn, "device_url") == "http://chosen.test"


def test_focus_order_limited_to_four(migrated):
    with contextlib.closing(db.connect(migrated.db_path)) as conn:
        conn.execute("INSERT INTO systems (beszel_id, beszel_name, focus_order) VALUES ('a','a',4)")
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                "INSERT INTO systems (beszel_id, beszel_name, focus_order) VALUES ('b','b',5)"
            )
        with pytest.raises(sqlite3.IntegrityError):  # one system per position
            conn.execute(
                "INSERT INTO systems (beszel_id, beszel_name, focus_order) VALUES ('c','c',4)"
            )


def test_push_state_keeps_last_good_on_error(migrated):
    with contextlib.closing(db.connect(migrated.db_path)) as conn:
        repo.record_push(conn, ok=True, frame_hash="abc")
        repo.record_push(conn, ok=False, frame_hash=None, error="timeout")
        s = repo.get_push_state(conn)
        assert s.last_hash == "abc"
        assert s.last_push_at is not None
        assert s.last_result == "error"
        assert s.last_error.endswith("timeout")


def test_backup_retention_and_protected_labels(migrated):
    for _ in range(4):
        backup.make_backup(migrated.db_path, migrated.backup_dir, keep=2)
    backup.make_backup(migrated.db_path, migrated.backup_dir, label="manual", keep=2)
    names = sorted(p.name for p in migrated.backup_dir.iterdir())
    assert len([n for n in names if n.endswith("-auto.db")]) == 2
    assert len([n for n in names if n.endswith("-manual.db")]) == 1
    assert backup.latest_age_hours(migrated.backup_dir) < 0.01


def test_backup_is_a_usable_database(migrated):
    path = backup.make_backup(migrated.db_path, migrated.backup_dir)
    with contextlib.closing(sqlite3.connect(path)) as conn:
        assert conn.execute("SELECT COUNT(*) FROM settings").fetchone()[0] >= 7
