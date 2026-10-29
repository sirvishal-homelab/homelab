"""Environment: infrastructure and secrets only (ADR-0005).

Everything users edit is in the database.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    data_dir: Path
    backup_keep: int
    backup_interval_hours: int
    web_concurrency: int
    terminus_url: str = ""
    terminus_login: str = ""
    terminus_password: str = field(default="", repr=False)  # a secret: never shown or logged
    public_base_url: str = ""  # the address *Terminus* uses to reach this app
    terminus_model_name: str = "og_plus"
    terminus_screen_name: str = "trmnl-calendar"
    terminus_verify_tls: bool = True
    push_interval_seconds: int = 900
    push_debounce_seconds: float = 5.0
    push_startup_delay_seconds: float = 10.0
    holiday_lookup: bool = True  # the optional online holiday lookup (ADR-0011)
    holiday_api_url: str = "https://date.nager.at/api/v3"

    @property
    def terminus_configured(self) -> bool:
        return all(
            (self.terminus_url, self.terminus_login, self.terminus_password, self.public_base_url)
        )

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
        terminus_url=e.get("TERMINUS_URL", "").strip().rstrip("/"),
        terminus_login=e.get("TERMINUS_LOGIN", "").strip(),
        terminus_password=e.get("TERMINUS_PASSWORD", ""),
        public_base_url=e.get("PUBLIC_BASE_URL", "").strip().rstrip("/"),
        terminus_model_name=e.get("TERMINUS_MODEL_NAME", "og_plus").strip() or "og_plus",
        terminus_screen_name=e.get("TERMINUS_SCREEN_NAME", "trmnl-calendar").strip()
        or "trmnl-calendar",
        terminus_verify_tls=e.get("TERMINUS_VERIFY_TLS", "true").strip().lower()
        not in ("0", "false", "no"),
        push_interval_seconds=max(60, int(e.get("TERMINUS_PUSH_INTERVAL", "900"))),
        push_debounce_seconds=float(e.get("PUSH_DEBOUNCE_SECONDS", "5")),
        push_startup_delay_seconds=float(e.get("PUSH_STARTUP_DELAY_SECONDS", "10")),
        holiday_lookup=e.get("HOLIDAY_LOOKUP", "on").strip().lower()
        not in ("0", "off", "false", "no"),
        holiday_api_url=e.get("HOLIDAY_API_URL", "").strip().rstrip("/")
        or "https://date.nager.at/api/v3",
    )
