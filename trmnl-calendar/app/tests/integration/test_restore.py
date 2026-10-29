import sqlite3

import pytest

from trmnl_calendar.storage import backup, db, migrate, restore
from trmnl_calendar.storage.repositories import events as repo


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def add(conn, title):
    return repo.create_event(conn, dict(title=title, section_id=1, start_date="2026-03-10"))


def titles(path):
    c = sqlite3.connect(path)
    try:
        return sorted(r[0] for r in c.execute("SELECT title FROM events"))
    finally:
        c.close()


def test_inspect_reports_contents(conn, settings):
    add(conn, "A")
    b = backup.make_backup(settings.db_path, settings.backup_dir, label="manual")
    info = restore.inspect(b)
    assert info["events"] == 1 and info["schema_version"] == migrate.current_version(conn)


def test_inspect_refuses_junk_and_missing(settings, tmp_path):
    junk = tmp_path / "calendar-x.db"
    junk.write_bytes(b"this is not sqlite" * 100)
    with pytest.raises(restore.RestoreError):
        restore.inspect(junk)
    with pytest.raises(restore.RestoreError):
        restore.inspect(tmp_path / "nope.db")
    other = tmp_path / "other.db"
    sqlite3.connect(other).execute("CREATE TABLE t (x)").connection.close()
    with pytest.raises(restore.RestoreError, match="not a calendar database"):
        restore.inspect(other)


def test_restore_replaces_data_and_keeps_the_old_one(conn, settings):
    add(conn, "Keep me")
    b = backup.make_backup(settings.db_path, settings.backup_dir, label="manual")
    add(conn, "Added later")
    conn.close()  # the app is stopped
    out = restore.restore(b, settings.db_path, settings.backup_dir)
    assert titles(settings.db_path) == ["Keep me"]
    assert titles(out["kept"]) == ["Added later", "Keep me"]
    assert "pre-restore" in out["kept"].name
    assert not (settings.db_path.parent / (settings.db_path.name + "-wal")).exists()


def test_restore_refuses_while_the_app_has_the_database_open(conn, settings):
    add(conn, "A")
    b = backup.make_backup(settings.db_path, settings.backup_dir, label="manual")
    holder = db.connect(settings.db_path)
    holder.execute("BEGIN IMMEDIATE")  # what a running request looks like
    try:
        with pytest.raises(restore.RestoreError, match="in use"):
            restore.restore(b, settings.db_path, settings.backup_dir)
    finally:
        holder.rollback()
        holder.close()
    assert titles(settings.db_path) == ["A"]


def test_restore_into_an_empty_volume(conn, settings, tmp_path):
    add(conn, "A")
    b = backup.make_backup(settings.db_path, settings.backup_dir, label="manual")
    fresh = tmp_path / "newvol" / "calendar.db"
    fresh.parent.mkdir()
    out = restore.restore(b, fresh, tmp_path / "newvol" / "backups")
    assert out["kept"] is None and titles(fresh) == ["A"]


def test_list_backups_marks_unusable_ones(conn, settings):
    good = backup.make_backup(settings.db_path, settings.backup_dir, label="manual")
    bad = settings.backup_dir / "calendar-00000000-000000-00-auto.db"
    bad.write_bytes(b"junk" * 50)
    got = dict(restore.list_backups(settings.backup_dir))
    assert isinstance(got[good], dict) and isinstance(got[bad], str)
