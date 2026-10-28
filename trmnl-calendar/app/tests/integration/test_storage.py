import sqlite3

from trmnl_calendar.storage import backup, db, migrate


def test_fresh_database_migrates_without_backup(settings):
    conn = db.connect(settings.db_path)
    applied = migrate.migrate(conn, settings.db_path, settings.backup_dir)
    assert applied == [name for _, name, _ in migrate.available()]
    assert migrate.current_version(conn) == len(applied)
    assert not settings.backup_dir.exists()
    assert migrate.migrate(conn, settings.db_path, settings.backup_dir) == []


def test_pragmas(settings):
    conn = db.connect(settings.db_path)
    assert conn.execute("PRAGMA journal_mode").fetchone()[0] == "wal"
    assert conn.execute("PRAGMA foreign_keys").fetchone()[0] == 1


def test_backup_is_consistent_and_pruned(settings):
    conn = db.connect(settings.db_path)
    migrate.migrate(conn, settings.db_path, settings.backup_dir)
    for _ in range(4):
        path = backup.make_backup(settings.db_path, settings.backup_dir, keep=2)
    assert len(list(settings.backup_dir.glob("*-auto*.db"))) == 2
    assert sqlite3.connect(path).execute("SELECT COUNT(*) FROM app_meta").fetchone()[0] == 1


def test_manual_backups_never_pruned(settings):
    conn = db.connect(settings.db_path)
    migrate.migrate(conn, settings.db_path, settings.backup_dir)
    for _ in range(3):
        backup.make_backup(settings.db_path, settings.backup_dir, label="manual", keep=0)
    assert len(list(settings.backup_dir.glob("*-manual*.db"))) == 3


def test_migration_backs_up_existing_db(settings, monkeypatch):
    conn = db.connect(settings.db_path)
    migrate.migrate(conn, settings.db_path, settings.backup_dir)
    real = migrate.available()
    n = len(real)
    extra = (n + 1, "9999_test.sql", "CREATE TABLE t (id INTEGER);")
    monkeypatch.setattr(migrate, "available", lambda: [*real, extra])
    assert migrate.migrate(conn, settings.db_path, settings.backup_dir) == ["9999_test.sql"]
    assert list(settings.backup_dir.glob(f"*pre-migration-v{n}.db"))
    assert migrate.current_version(conn) == n + 1


def test_failed_migration_rolls_back(settings, monkeypatch):
    conn = db.connect(settings.db_path)
    migrate.migrate(conn, settings.db_path, settings.backup_dir)
    real = migrate.available()
    n = len(real)
    bad = (n + 1, "9998_bad.sql", "CREATE TABLE ok (id INTEGER); CREATE TABLE ok (id INTEGER);")
    monkeypatch.setattr(migrate, "available", lambda: [*real, bad])
    try:
        migrate.migrate(conn, settings.db_path, settings.backup_dir)
    except sqlite3.Error:
        pass
    assert migrate.current_version(conn) == n
    assert conn.execute("SELECT name FROM sqlite_master WHERE name='ok'").fetchone() is None
