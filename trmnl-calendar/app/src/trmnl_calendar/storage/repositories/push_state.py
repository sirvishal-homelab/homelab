"""The publisher's persistent state: one row per display (docs/03)."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass

from trmnl_calendar.storage.repositories.views import DISPLAY_ID


@dataclass(frozen=True)
class PushState:
    screen_id: int | None
    last_hash: str | None
    last_status: str
    last_error: str | None
    last_attempt_at: str | None
    last_pushed_at: str | None
    failures: int


def get_state(conn: sqlite3.Connection) -> PushState:
    r = conn.execute("SELECT * FROM push_state WHERE display_id = ?", (DISPLAY_ID,)).fetchone()
    return PushState(
        r["screen_id"],
        r["last_hash"],
        r["last_status"],
        r["last_error"],
        r["last_attempt_at"],
        r["last_pushed_at"],
        r["failures"],
    )


def save(conn: sqlite3.Connection, **fields) -> None:
    allowed = {
        "screen_id",
        "last_hash",
        "last_status",
        "last_error",
        "last_attempt_at",
        "last_pushed_at",
        "failures",
    }
    bad = set(fields) - allowed
    if bad:
        raise ValueError(f"unknown push_state fields: {sorted(bad)}")
    if not fields:
        return
    sets = ", ".join(f"{k} = ?" for k in fields)
    conn.execute(
        f"UPDATE push_state SET {sets} WHERE display_id = ?",  # noqa: S608 (keys are whitelisted)
        (*fields.values(), DISPLAY_ID),
    )
    conn.commit()


def remember_device(conn: sqlite3.Connection, devices: list[dict], model_id: int | None) -> None:
    """Keep the device refresh time (seconds), so the status text can quote Terminus's setting."""
    mine = [d for d in devices if model_id is not None and d.get("model_id") == model_id]
    device = (mine or devices or [None])[0]
    seconds = device.get("refresh_rate") if device else None
    if isinstance(seconds, int) and not isinstance(seconds, bool) and seconds > 0:
        conn.execute(
            "INSERT INTO app_meta (key, value) VALUES ('device_refresh_seconds', ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (str(seconds),),
        )
        conn.commit()


def device_refresh_seconds(conn: sqlite3.Connection) -> int | None:
    r = conn.execute("SELECT value FROM app_meta WHERE key = 'device_refresh_seconds'").fetchone()
    return int(r["value"]) if r and str(r["value"]).isdigit() else None
