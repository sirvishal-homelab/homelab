"""Environment: secrets and bootstrap values only (ADR-0003).

Everything the owner tunes (thresholds, timing, focus systems) lives in the database and is
edited on the settings page. `SCREEN_URL` only seeds the screen address on first start.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    data_dir: Path
    backup_keep: int = 7
    backup_interval_hours: int = 24
    web_concurrency: int = 1
    screen_url: str = ""  # seed for the `device_url` setting
    beszel_url: str = ""
    beszel_user: str = ""
    beszel_password: str = field(default="", repr=False)  # a secret: never shown or logged
    timezone: str = "America/Halifax"
    focus_systems: tuple[str, ...] = ()  # until the settings page (S-4); names or Beszel ids

    @property
    def db_path(self) -> Path:
        return self.data_dir / "screen.db"

    @property
    def backup_dir(self) -> Path:
        return self.data_dir / "backups"

    @property
    def beszel_configured(self) -> bool:
        return all((self.beszel_url, self.beszel_user, self.beszel_password))


def _url(value: str) -> str:
    return value.strip().rstrip("/")


def load_settings(env: dict[str, str] | None = None) -> Settings:
    e = os.environ if env is None else env
    return Settings(
        data_dir=Path(e.get("DATA_DIR", "/data")),
        backup_keep=max(1, int(e.get("BACKUP_KEEP", "7"))),
        backup_interval_hours=max(1, int(e.get("BACKUP_INTERVAL_HOURS", "24"))),
        web_concurrency=int(e.get("WEB_CONCURRENCY", "1")),
        screen_url=_url(e.get("SCREEN_URL", "")),
        beszel_url=_url(e.get("BESZEL_URL", "")),
        beszel_user=e.get("BESZEL_USER", "").strip(),
        beszel_password=e.get("BESZEL_PASSWORD", ""),
        timezone=e.get("TZ", "America/Halifax").strip() or "America/Halifax",
        focus_systems=tuple(n.strip() for n in e.get("FOCUS_SYSTEMS", "").split(",") if n.strip())[
            :4
        ],
    )
