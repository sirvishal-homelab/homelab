"""Idempotency keys for creates (docs/06): a retry returns the stored response."""

from __future__ import annotations

import json
import sqlite3
from datetime import UTC, datetime, timedelta

TTL_HOURS = 24


def _stamp(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def lookup(conn: sqlite3.Connection, key: str, method_and_path: str) -> tuple[int, dict] | None:
    cutoff = _stamp(datetime.now(UTC) - timedelta(hours=TTL_HOURS))
    row = conn.execute(
        "SELECT status, response_json FROM idempotency_keys "
        "WHERE key = ? AND method_and_path = ? AND created_at >= ?",
        (key, method_and_path, cutoff),
    ).fetchone()
    return (row["status"], json.loads(row["response_json"])) if row else None


def store(
    conn: sqlite3.Connection, key: str, method_and_path: str, status: int, body: dict
) -> None:
    conn.execute(
        "INSERT OR REPLACE INTO idempotency_keys VALUES (?, ?, ?, ?, ?)",
        (key, method_and_path, status, json.dumps(body), _stamp(datetime.now(UTC))),
    )
    conn.commit()


def purge(conn: sqlite3.Connection) -> int:
    cutoff = _stamp(datetime.now(UTC) - timedelta(hours=TTL_HOURS))
    cur = conn.execute("DELETE FROM idempotency_keys WHERE created_at < ?", (cutoff,))
    conn.commit()
    return cur.rowcount
