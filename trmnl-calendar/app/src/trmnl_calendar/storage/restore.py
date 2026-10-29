"""Check and restore a backup (docs/19 "Restore drill").

Run with the app stopped:  python -m trmnl_calendar.storage.restore --list
                           python -m trmnl_calendar.storage.restore <backup-file>
"""

from __future__ import annotations

import argparse
import sqlite3
import sys
from datetime import UTC, datetime
from pathlib import Path

from trmnl_calendar.config import load_settings
from trmnl_calendar.storage import backup, migrate


class RestoreError(Exception):
    """The backup cannot be used, or the app is still running. The message says what to do."""


def inspect(path: Path) -> dict:
    """Open a backup read-only and report what is in it. Raises RestoreError if unusable."""
    if not path.is_file():
        raise RestoreError(f"{path} does not exist.")
    try:
        conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    except sqlite3.Error as exc:
        raise RestoreError(f"{path.name} cannot be opened: {exc}") from exc
    try:
        try:
            check = conn.execute("PRAGMA integrity_check").fetchall()
            if check != [("ok",)]:
                raise RestoreError(f"{path.name} failed the integrity check: {check[:3]}")
            version = migrate.current_version(conn)
            events = conn.execute(
                "SELECT COUNT(*) FROM events WHERE deleted_at IS NULL"
            ).fetchone()[0]
            people = conn.execute("SELECT COUNT(*) FROM people").fetchone()[0]
        except sqlite3.Error as exc:
            raise RestoreError(f"{path.name} is not a calendar database: {exc}") from exc
    finally:
        conn.close()
    newest = max((n for n, _, _ in migrate.available()), default=0)
    if version > newest:
        raise RestoreError(
            f"{path.name} is from a newer version of the app (schema {version}, this one knows "
            f"{newest}). Upgrade the app first."
        )
    return {"schema_version": version, "events": events, "people": people}


def _require_stopped(db_path: Path) -> None:
    """A write lock on the live database fails while the app is using it."""
    if not db_path.exists():
        return
    conn = sqlite3.connect(db_path, timeout=0.5)
    try:
        conn.execute("BEGIN EXCLUSIVE")
        conn.rollback()
    except sqlite3.OperationalError as exc:
        raise RestoreError(
            "The database is in use. Stop the app first: `docker compose stop calendar`."
        ) from exc
    finally:
        conn.close()


def restore(backup_file: Path, db_path: Path, backup_dir: Path) -> dict:
    """Replace the live database with `backup_file`, keeping the old one as a pre-restore backup."""
    info = inspect(backup_file)
    _require_stopped(db_path)
    kept = None
    if db_path.exists():
        kept = backup.make_backup(db_path, backup_dir, label="pre-restore")
    tmp = db_path.with_name(db_path.name + ".restoring")
    tmp.unlink(missing_ok=True)
    src = sqlite3.connect(f"file:{backup_file}?mode=ro", uri=True)
    try:
        dst = sqlite3.connect(tmp)
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    for suffix in ("-wal", "-shm"):
        Path(str(db_path) + suffix).unlink(missing_ok=True)
    tmp.replace(db_path)
    return {**info, "kept": kept}


def list_backups(backup_dir: Path) -> list[tuple[Path, dict | str]]:
    out: list[tuple[Path, dict | str]] = []
    for f in sorted(backup_dir.glob("calendar-*.db"), reverse=True):
        try:
            out.append((f, inspect(f)))
        except RestoreError as exc:
            out.append((f, str(exc)))
    return out


def _age(f: Path) -> str:
    t = datetime.fromtimestamp(f.stat().st_mtime, UTC)
    return t.strftime("%Y-%m-%d %H:%M UTC")


def main() -> int:
    p = argparse.ArgumentParser(description="List, check or restore calendar backups.")
    p.add_argument("file", nargs="?", help="backup file to restore (a path, or a name in backups/)")
    p.add_argument("--list", action="store_true", help="list backups and what is in them")
    a = p.parse_args()
    s = load_settings()
    if a.list or not a.file:
        for f, info in list_backups(s.backup_dir):
            what = (
                f"{info['events']} events, {info['people']} people, schema {info['schema_version']}"
                if isinstance(info, dict)
                else f"UNUSABLE: {info}"
            )
            print(f"{f.name}  {_age(f)}  {what}")
        return 0
    path = Path(a.file)
    if not path.is_file():
        path = s.backup_dir / a.file
    try:
        out = restore(path, s.db_path, s.backup_dir)
    except RestoreError as exc:
        print(f"Not restored: {exc}", file=sys.stderr)
        return 1
    print(
        f"Restored {path.name}: {out['events']} events, {out['people']} people. "
        f"The previous database was saved as {out['kept'].name if out['kept'] else '(none)'}. "
        "Start the app again: `docker compose up -d`."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
