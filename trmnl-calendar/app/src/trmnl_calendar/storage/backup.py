"""Consistent SQLite backups (online backup API, safe while the app runs) with retention."""

from __future__ import annotations

import argparse
import sqlite3
import time
from datetime import UTC, datetime
from pathlib import Path

from trmnl_calendar.config import load_settings


def make_backup(db_path: Path, backup_dir: Path, *, label: str = "auto", keep: int = 14) -> Path:
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(UTC).strftime("%Y%m%d-%H%M%S%f")
    # name = time, then a two-digit counter, then the label, so sorting by name sorts by age
    n = 0
    dest = backup_dir / f"calendar-{stamp}-{n:02d}-{label}.db"
    while dest.exists():
        n += 1
        dest = backup_dir / f"calendar-{stamp}-{n:02d}-{label}.db"
    src = sqlite3.connect(db_path)
    try:
        dst = sqlite3.connect(dest)
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    prune(backup_dir, keep, protect=dest)
    return dest


def prune(backup_dir: Path, keep: int, protect: Path | None = None) -> None:
    """Keep the newest `keep` automatic backups; pre-migration and manual ones are never pruned."""
    autos = sorted(backup_dir.glob("calendar-*-auto.db"))
    for old in autos[: max(0, len(autos) - keep)]:
        if old != protect:
            old.unlink()


def latest_age_hours(backup_dir: Path) -> float | None:
    files = list(backup_dir.glob("calendar-*.db"))
    if not files:
        return None
    return (time.time() - max(f.stat().st_mtime for f in files)) / 3600


def main() -> None:
    p = argparse.ArgumentParser(description="Back up the calendar database now.")
    p.add_argument("--now", action="store_true", required=True)
    p.add_argument("--label", default="manual")
    a = p.parse_args()
    s = load_settings()
    print(make_backup(s.db_path, s.backup_dir, label=a.label, keep=s.backup_keep))


if __name__ == "__main__":
    main()
