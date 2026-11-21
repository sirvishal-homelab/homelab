"""Stored flight status and the monthly call counter (migration 0009, docs/21)."""

from __future__ import annotations

import sqlite3
from collections.abc import Iterable
from datetime import UTC, date, datetime, timedelta

from trmnl_calendar.domain.flights import FlightInfo

KEEP_DAYS = 7  # the provider's caching limit (AeroDataBox Basic plan)
_FMT = "%Y-%m-%dT%H:%M:%SZ"


def _t(value: datetime | None) -> str | None:
    return value.astimezone(UTC).strftime(_FMT) if value else None


def _p(value: str | None) -> datetime | None:
    return datetime.strptime(value, _FMT).replace(tzinfo=UTC) if value else None


def _info(r: sqlite3.Row) -> FlightInfo:
    return FlightInfo(
        code=r["code"],
        dep_date=date.fromisoformat(r["dep_date"]),
        found=bool(r["found"]),
        status=r["status"],
        sched_dep=_p(r["sched_dep"]),
        revised_dep=_p(r["revised_dep"]),
        sched_arr=_p(r["sched_arr"]),
        revised_arr=_p(r["revised_arr"]),
        sched_dep_local=r["sched_dep_local"],
        revised_dep_local=r["revised_dep_local"],
        dep_airport=r["dep_airport"],
        arr_airport=r["arr_airport"],
        dep_gate=r["dep_gate"],
        sched_arr_local=r["sched_arr_local"],
        revised_arr_local=r["revised_arr_local"],
        dep_icao=r["dep_icao"],
        arr_icao=r["arr_icao"],
    )


def get(conn: sqlite3.Connection, code: str, day: date) -> sqlite3.Row | None:
    return conn.execute(
        "SELECT * FROM flight_status WHERE code = ? AND dep_date = ?", (code, day.isoformat())
    ).fetchone()


def info(row: sqlite3.Row | None) -> FlightInfo | None:
    return _info(row) if row is not None else None


def infos_for(
    conn: sqlite3.Connection, keys: Iterable[tuple[str, date]]
) -> dict[tuple[str, date], FlightInfo]:
    """Stored answers for these (code, day) pairs; missing pairs are simply absent."""
    keys = set(keys)
    if not keys:
        return {}
    days = sorted({d.isoformat() for _, d in keys})
    marks = ",".join("?" * len(days))
    rows = conn.execute(f"SELECT * FROM flight_status WHERE dep_date IN ({marks})", days).fetchall()
    out = {}
    for r in rows:
        key = (r["code"], date.fromisoformat(r["dep_date"]))
        if key in keys:
            out[key] = _info(r)
    return out


def save(
    conn: sqlite3.Connection, info: FlightInfo, checked_at: datetime, next_due: datetime | None
) -> None:
    conn.execute(
        """INSERT INTO flight_status (code, dep_date, found, status, sched_dep, revised_dep,
               sched_arr, revised_arr, sched_dep_local, revised_dep_local, dep_airport,
               arr_airport, dep_gate, sched_arr_local, revised_arr_local, dep_icao, arr_icao,
               checked_at, next_due_at, last_error)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, '')
           ON CONFLICT (code, dep_date) DO UPDATE SET
               found = excluded.found, status = excluded.status,
               sched_dep = excluded.sched_dep, revised_dep = excluded.revised_dep,
               sched_arr = excluded.sched_arr, revised_arr = excluded.revised_arr,
               sched_dep_local = excluded.sched_dep_local,
               revised_dep_local = excluded.revised_dep_local,
               dep_airport = excluded.dep_airport, arr_airport = excluded.arr_airport,
               dep_gate = excluded.dep_gate, sched_arr_local = excluded.sched_arr_local,
               revised_arr_local = excluded.revised_arr_local, dep_icao = excluded.dep_icao,
               arr_icao = excluded.arr_icao, checked_at = excluded.checked_at,
               next_due_at = excluded.next_due_at, last_error = ''""",
        (
            info.code,
            info.dep_date.isoformat(),
            int(info.found),
            info.status,
            _t(info.sched_dep),
            _t(info.revised_dep),
            _t(info.sched_arr),
            _t(info.revised_arr),
            info.sched_dep_local,
            info.revised_dep_local,
            info.dep_airport,
            info.arr_airport,
            info.dep_gate,
            info.sched_arr_local,
            info.revised_arr_local,
            info.dep_icao,
            info.arr_icao,
            _t(checked_at),
            _t(next_due),
        ),
    )
    conn.commit()


def save_error(
    conn: sqlite3.Connection,
    code: str,
    day: date,
    message: str,
    checked_at: datetime,
    retry_at: datetime,
) -> None:
    """Keep the last good answer; note the failure and when to try again."""
    conn.execute(
        """INSERT INTO flight_status (code, dep_date, found, checked_at, next_due_at, last_error)
           VALUES (?, ?, 0, ?, ?, ?)
           ON CONFLICT (code, dep_date) DO UPDATE SET checked_at = excluded.checked_at,
               next_due_at = excluded.next_due_at, last_error = excluded.last_error""",
        (code, day.isoformat(), _t(checked_at), _t(retry_at), message[:300]),
    )
    conn.commit()


def is_due(row: sqlite3.Row, now: datetime) -> bool:
    due = _p(row["next_due_at"])
    return due is not None and due <= now


def checked_at(row: sqlite3.Row) -> datetime | None:
    return _p(row["checked_at"])


def calls_this_month(conn: sqlite3.Connection, now: datetime) -> int:
    row = conn.execute(
        "SELECT count FROM flight_calls WHERE month = ?", (now.strftime("%Y-%m"),)
    ).fetchone()
    return row["count"] if row else 0


def count_call(conn: sqlite3.Connection, now: datetime) -> None:
    conn.execute(
        "INSERT INTO flight_calls (month, count) VALUES (?, 1) "
        "ON CONFLICT (month) DO UPDATE SET count = count + 1",
        (now.strftime("%Y-%m"),),
    )
    conn.commit()


def purge(conn: sqlite3.Connection, now: datetime) -> int:
    """Delete answers older than the provider allows us to keep."""
    cur = conn.execute(
        "DELETE FROM flight_status WHERE checked_at < ?", (_t(now - timedelta(days=KEEP_DAYS)),)
    )
    conn.commit()
    return cur.rowcount
