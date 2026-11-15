"""Small typed helpers over the settings and push_state tables."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any


def get_setting(conn: sqlite3.Connection, key: str, default: Any = None) -> Any:
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    return default if row is None else json.loads(row["value"])


def set_setting(conn: sqlite3.Connection, key: str, value: Any) -> None:
    conn.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, json.dumps(value)),
    )
    conn.commit()


def seed_device_url(conn: sqlite3.Connection, screen_url: str) -> str:
    """Use SCREEN_URL from the environment only if the setting is still empty."""
    current = get_setting(conn, "device_url", "")
    if not current and screen_url:
        set_setting(conn, "device_url", screen_url)
        return screen_url
    return current


@dataclass(frozen=True)
class PushState:
    last_hash: str | None
    last_push_at: str | None
    last_result: str | None
    last_error: str | None


def get_push_state(conn: sqlite3.Connection) -> PushState:
    r = conn.execute(
        "SELECT last_hash, last_push_at, last_result, last_error FROM push_state WHERE id = 1"
    ).fetchone()
    return PushState(r["last_hash"], r["last_push_at"], r["last_result"], r["last_error"])


def record_push(
    conn: sqlite3.Connection, *, ok: bool, frame_hash: str | None, error: str | None = None
) -> None:
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    if ok:
        conn.execute(
            "UPDATE push_state SET last_hash = ?, last_push_at = ?, last_result = 'ok', "
            "last_error = NULL WHERE id = 1",
            (frame_hash, now),
        )
    else:  # keep the last good hash and time; record what went wrong
        conn.execute(
            "UPDATE push_state SET last_result = 'error', last_error = ? WHERE id = 1",
            (f"{now} {error}",),
        )
    conn.commit()


# --- container memory (docs/05 section 4) -----------------------------------------------------


def load_container_memory(
    conn: sqlite3.Connection,
) -> tuple[dict[tuple[str, str], int], frozenset[tuple[str, str]]]:
    rows = conn.execute(
        "SELECT system_id, name, missing_polls, allowed_to_stop FROM known_containers"
    ).fetchall()
    memory = {(r["system_id"], r["name"]): r["missing_polls"] for r in rows}
    allowed = frozenset((r["system_id"], r["name"]) for r in rows if r["allowed_to_stop"])
    return memory, allowed


def save_container_memory(conn: sqlite3.Connection, memory: dict[tuple[str, str], int]) -> None:
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn.executemany(
        "INSERT INTO known_containers (system_id, name, missing_polls, last_running) "
        "VALUES (?, ?, ?, CASE WHEN ? = 0 THEN ? END) "
        "ON CONFLICT(system_id, name) DO UPDATE SET missing_polls = excluded.missing_polls, "
        "last_running = COALESCE(excluded.last_running, known_containers.last_running)",
        [(sid, name, n, n, now) for (sid, name), n in memory.items()],
    )
    conn.commit()
