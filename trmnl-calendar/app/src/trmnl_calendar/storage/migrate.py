"""Forward-only numbered migrations with a backup before any change to an existing database."""

from __future__ import annotations

import re
import sqlite3
from importlib import resources
from pathlib import Path

from trmnl_calendar.storage.backup import make_backup

_NAME = re.compile(r"^(\d{4})_[a-z0-9_]+\.sql$")


def available() -> list[tuple[int, str, str]]:
    root = resources.files("trmnl_calendar.storage") / "migrations"
    found = []
    for f in root.iterdir():
        m = _NAME.match(f.name)
        if m:
            found.append((int(m.group(1)), f.name, f.read_text(encoding="utf-8")))
    found.sort()
    versions = [v for v, _, _ in found]
    if versions != list(range(1, len(versions) + 1)):
        raise RuntimeError(f"migration numbers must be 1..N without gaps, got {versions}")
    return found


def current_version(conn: sqlite3.Connection) -> int:
    conn.execute(
        "CREATE TABLE IF NOT EXISTS schema_migrations "
        "(version INTEGER PRIMARY KEY, name TEXT NOT NULL, applied_at TEXT NOT NULL "
        "DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now')))"
    )
    row = conn.execute("SELECT MAX(version) FROM schema_migrations").fetchone()
    return row[0] or 0


def migrate(conn: sqlite3.Connection, db_path: Path, backup_dir: Path) -> list[str]:
    """Apply pending migrations. Returns the names applied."""
    have = current_version(conn)
    pending = [m for m in available() if m[0] > have]
    if not pending:
        return []
    if have > 0:  # an existing database: back it up before changing it
        make_backup(db_path, backup_dir, label=f"pre-migration-v{have}")
    applied = []
    for version, name, sql in pending:
        try:
            conn.executescript(
                f"BEGIN;\n{sql}\nINSERT INTO schema_migrations (version, name) "
                f"VALUES ({version}, '{name}');\nCOMMIT;"
            )
        except sqlite3.Error:
            if conn.in_transaction:
                conn.execute("ROLLBACK")
            raise
        applied.append(name)
    return applied
