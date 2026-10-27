"""Environment: infrastructure and secrets only (ADR-0005).

Everything users edit is in the database.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    data_dir: Path
    backup_keep: int
    backup_interval_hours: int
    web_concurrency: int

    @property
    def db_path(self) -> Path:
        return self.data_dir / "calendar.db"

    @property
    def backup_dir(self) -> Path:
        return self.data_dir / "backups"


def load_settings(env: dict[str, str] | None = None) -> Settings:
    e = os.environ if env is None else env
    return Settings(
        data_dir=Path(e.get("DATA_DIR", "/data")),
        backup_keep=int(e.get("BACKUP_KEEP", "14")),
        backup_interval_hours=int(e.get("BACKUP_INTERVAL_HOURS", "24")),
        web_concurrency=int(e.get("WEB_CONCURRENCY", "1")),
    )
