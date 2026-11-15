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


# --- systems and preferences (S-4 settings page) ---------------------------------------------


@dataclass(frozen=True)
class SystemPrefs:
    beszel_id: str
    beszel_name: str
    display_name: str | None
    focus_order: int | None
    ignore_alerts: bool
    last_seen: str


def sync_systems(conn: sqlite3.Connection, systems: list[tuple[str, str]]) -> None:
    """Record every system Beszel reports (id, name). New ones start unfocused."""
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn.executemany(
        "INSERT INTO systems (beszel_id, beszel_name, last_seen) VALUES (?, ?, ?) "
        "ON CONFLICT(beszel_id) DO UPDATE SET beszel_name = excluded.beszel_name, "
        "last_seen = excluded.last_seen",
        [(sid, name, now) for sid, name in systems],
    )
    conn.commit()


def list_systems(conn: sqlite3.Connection) -> list[SystemPrefs]:
    rows = conn.execute(
        "SELECT beszel_id, beszel_name, display_name, focus_order, ignore_alerts, last_seen "
        "FROM systems ORDER BY focus_order IS NULL, focus_order, lower(beszel_name)"
    ).fetchall()
    return [
        SystemPrefs(
            r["beszel_id"],
            r["beszel_name"],
            r["display_name"],
            r["focus_order"],
            bool(r["ignore_alerts"]),
            r["last_seen"],
        )
        for r in rows
    ]


def focus_ids(conn: sqlite3.Connection) -> list[str]:
    return [p.beszel_id for p in list_systems(conn) if p.focus_order is not None]


def labels(conn: sqlite3.Connection) -> dict[str, str]:
    """Label per system id: the display name if set, else Beszel's name."""
    return {p.beszel_id: p.display_name or p.beszel_name for p in list_systems(conn)}


def ignored_systems(conn: sqlite3.Connection) -> frozenset[str]:
    return frozenset(p.beszel_id for p in list_systems(conn) if p.ignore_alerts)


def save_system_prefs(
    conn: sqlite3.Connection, prefs: dict[str, tuple[int | None, str | None, bool]]
) -> None:
    """prefs: beszel_id -> (focus_order, display_name, ignore_alerts). One transaction."""
    with conn:
        conn.execute("UPDATE systems SET focus_order = NULL")  # free the unique positions
        conn.executemany(
            "UPDATE systems SET focus_order = ?, display_name = ?, ignore_alerts = ? "
            "WHERE beszel_id = ?",
            [(o, n, int(i), sid) for sid, (o, n, i) in prefs.items()],
        )


def seed_focus_by_name(conn: sqlite3.Connection, names: tuple[str, ...]) -> bool:
    """One-time move of FOCUS_SYSTEMS (.env, S-2) into the database. Returns True if done."""
    if get_setting(conn, "focus_seeded", False) or not names:
        return False
    if focus_ids(conn):
        set_setting(conn, "focus_seeded", True)
        return False
    by_name = {p.beszel_name.lower(): p.beszel_id for p in list_systems(conn)}
    by_id = {p.beszel_id for p in list_systems(conn)}
    picked = []
    for n in names:
        sid = n if n in by_id else by_name.get(n.lower())
        if sid and sid not in picked:
            picked.append(sid)
    if not picked:
        return False  # systems not known yet: try again after the next read
    with conn:
        for pos, sid in enumerate(picked[:4], start=1):
            conn.execute("UPDATE systems SET focus_order = ? WHERE beszel_id = ?", (pos, sid))
    set_setting(conn, "focus_seeded", True)
    return True


@dataclass(frozen=True)
class KnownContainer:
    system_id: str
    system_name: str
    name: str
    allowed_to_stop: bool
    missing_polls: int
    last_running: str | None


def list_containers(conn: sqlite3.Connection) -> list[KnownContainer]:
    rows = conn.execute(
        "SELECT k.system_id, COALESCE(s.display_name, s.beszel_name, k.system_id) AS sname, "
        "k.name, k.allowed_to_stop, k.missing_polls, k.last_running "
        "FROM known_containers k LEFT JOIN systems s ON s.beszel_id = k.system_id "
        "ORDER BY lower(sname), lower(k.name)"
    ).fetchall()
    return [
        KnownContainer(
            r["system_id"],
            r["sname"],
            r["name"],
            bool(r["allowed_to_stop"]),
            r["missing_polls"],
            r["last_running"],
        )
        for r in rows
    ]


def save_allowed_to_stop(conn: sqlite3.Connection, allowed: set[tuple[str, str]]) -> None:
    with conn:
        conn.execute("UPDATE known_containers SET allowed_to_stop = 0")
        conn.executemany(
            "UPDATE known_containers SET allowed_to_stop = 1 WHERE system_id = ? AND name = ?",
            sorted(allowed),
        )


def forget_container(conn: sqlite3.Connection, system_id: str, name: str) -> bool:
    with conn:
        cur = conn.execute(
            "DELETE FROM known_containers WHERE system_id = ? AND name = ?", (system_id, name)
        )
    return cur.rowcount > 0


def clear_push_hash(conn: sqlite3.Connection) -> None:
    """Make the next cycle upload even if the picture did not change (Push now)."""
    conn.execute("UPDATE push_state SET last_hash = NULL WHERE id = 1")
    conn.commit()
