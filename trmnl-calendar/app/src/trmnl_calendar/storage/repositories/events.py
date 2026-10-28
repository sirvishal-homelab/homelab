"""Event persistence: create, read, patch with revision check, soft delete, restore, skips."""

from __future__ import annotations

import sqlite3
import uuid
from datetime import UTC, date, datetime, timedelta

from trmnl_calendar.domain.models import Event
from trmnl_calendar.domain.validation import parse_date, validate_event
from trmnl_calendar.errors import EditConflict, NotFound, ValidationFailed
from trmnl_calendar.storage.repositories.calendar import CALENDAR_ID

RECOVERY_DAYS = 30
_FIELDS = (
    "title",
    "location",
    "section_id",
    "start_date",
    "end_date",
    "start_time",
    "end_time",
    "freq",
    "interval",
    "by_weekday",
    "month_mode",
    "month_nth",
    "until",
)


def _now() -> str:
    return datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def row_to_event(r: sqlite3.Row) -> Event:
    return Event(
        id=r["id"],
        uid=r["uid"],
        title=r["title"],
        location=r["location"],
        section_id=r["section_id"],
        start_date=date.fromisoformat(r["start_date"]),
        end_date=date.fromisoformat(r["end_date"]) if r["end_date"] else None,
        start_time=r["start_time"],
        end_time=r["end_time"],
        freq=r["freq"],
        interval=r["interval"],
        by_weekday=tuple(c for c in r["by_weekday"].split(",") if c),
        month_mode=r["month_mode"],
        month_nth=r["month_nth"],
        until=date.fromisoformat(r["until"]) if r["until"] else None,
        rev=r["rev"],
    )


def event_to_form(ev: Event) -> dict:
    """The values validate_event expects (dates as ISO text), for merging patches."""
    return {
        "title": ev.title,
        "location": ev.location,
        "section_id": ev.section_id,
        "start_date": ev.start_date.isoformat(),
        "end_date": ev.end_date.isoformat() if ev.end_date else None,
        "start_time": ev.start_time,
        "end_time": ev.end_time,
        "freq": ev.freq,
        "interval": ev.interval,
        "by_weekday": list(ev.by_weekday),
        "month_mode": ev.month_mode,
        "month_nth": ev.month_nth,
        "until": ev.until.isoformat() if ev.until else None,
    }


def _check_section(conn: sqlite3.Connection, section_id: int | None) -> None:
    ok = conn.execute(
        "SELECT 1 FROM sections WHERE id = ? AND calendar_id = ?", (section_id, CALENDAR_ID)
    ).fetchone()
    if not ok:
        raise ValidationFailed({"section_id": "Choose a category"})


def _params(clean: dict) -> dict:
    p = dict(clean)
    for k in ("start_date", "end_date", "until"):
        p[k] = p[k].isoformat() if p[k] else None
    p["by_weekday"] = ",".join(p["by_weekday"])
    return p


def create_event(conn: sqlite3.Connection, data: dict, *, person_id: int | None = None) -> Event:
    clean, errors = validate_event(data)
    if errors:
        raise ValidationFailed(errors)
    _check_section(conn, clean["section_id"])
    now = _now()
    cur = conn.execute(
        """INSERT INTO events (uid, calendar_id, section_id, title, location, start_date, end_date,
               start_time, end_time, freq, interval, by_weekday, month_mode, month_nth, until,
               rev, updated_by_person_id, created_at, updated_at)
           VALUES (:uid, :cal, :section_id, :title, :location, :start_date, :end_date,
               :start_time, :end_time, :freq, :interval, :by_weekday, :month_mode, :month_nth,
               :until, 1, :person, :now, :now)""",
        {
            **_params(clean),
            "uid": str(uuid.uuid4()),
            "cal": CALENDAR_ID,
            "person": person_id,
            "now": now,
        },
    )
    conn.commit()
    return get_event(conn, cur.lastrowid)


def get_event(conn: sqlite3.Connection, event_id: int, *, include_deleted: bool = False) -> Event:
    sql = "SELECT * FROM events WHERE id = ? AND calendar_id = ?"
    if not include_deleted:
        sql += " AND deleted_at IS NULL"
    row = conn.execute(sql, (event_id, CALENDAR_ID)).fetchone()
    if not row:
        raise NotFound("That event no longer exists.")
    return row_to_event(row)


def meta(conn: sqlite3.Connection, event_id: int) -> dict:
    r = conn.execute(
        "SELECT created_at, updated_at, deleted_at, updated_by_person_id FROM events WHERE id = ?",
        (event_id,),
    ).fetchone()
    return dict(r) if r else {}


def list_events(conn: sqlite3.Connection, window_start: date, window_end: date) -> list[Event]:
    """Series that could produce an occurrence in the window (not expanded)."""
    rows = conn.execute(
        """SELECT * FROM events
           WHERE calendar_id = ? AND deleted_at IS NULL AND start_date <= ?
             AND (freq != 'none' OR COALESCE(end_date, start_date) >= ?)
             AND (until IS NULL OR until >= ? OR COALESCE(end_date, start_date) >= ?)
           ORDER BY start_date, id""",
        (
            CALENDAR_ID,
            window_end.isoformat(),
            window_start.isoformat(),
            window_start.isoformat(),
            window_start.isoformat(),
        ),
    ).fetchall()
    return [row_to_event(r) for r in rows]


def list_upcoming(conn: sqlite3.Connection, today: date, limit: int = 200) -> list[Event]:
    """Series still relevant on or after today: repeating without a passed end, or not yet over."""
    rows = conn.execute(
        """SELECT * FROM events
           WHERE calendar_id = ? AND deleted_at IS NULL
             AND ((freq = 'none' AND COALESCE(end_date, start_date) >= ?)
               OR (freq != 'none' AND (until IS NULL OR until >= ?)))
           ORDER BY start_date, id LIMIT ?""",
        (CALENDAR_ID, today.isoformat(), today.isoformat(), limit),
    ).fetchall()
    return [row_to_event(r) for r in rows]


def update_event(
    conn: sqlite3.Connection,
    event_id: int,
    patch: dict,
    expected_rev: int,
    *,
    person_id: int | None = None,
) -> Event:
    current = get_event(conn, event_id)
    if current.rev != expected_rev:
        raise EditConflict(
            "This event was changed by someone else.",
            extra={"current_rev": current.rev, "changed_at": meta(conn, event_id)["updated_at"]},
        )
    unknown = set(patch) - set(_FIELDS)
    if unknown:
        raise ValidationFailed({k: "Unknown field" for k in sorted(unknown)})
    merged = {**event_to_form(current), **patch}
    clean, errors = validate_event(merged)
    if errors:
        raise ValidationFailed(errors)
    _check_section(conn, clean["section_id"])
    cur = conn.execute(
        """UPDATE events SET section_id=:section_id, title=:title, location=:location,
               start_date=:start_date, end_date=:end_date, start_time=:start_time,
               end_time=:end_time, freq=:freq, interval=:interval, by_weekday=:by_weekday,
               month_mode=:month_mode, month_nth=:month_nth, until=:until,
               rev = rev + 1, updated_by_person_id=:person, updated_at=:now
           WHERE id=:id AND rev=:rev AND deleted_at IS NULL""",
        {**_params(clean), "person": person_id, "now": _now(), "id": event_id, "rev": expected_rev},
    )
    conn.commit()
    if cur.rowcount != 1:  # lost a race between the check and the write
        raise EditConflict("This event was changed by someone else.")
    return get_event(conn, event_id)


def soft_delete(conn: sqlite3.Connection, event_id: int, expected_rev: int) -> Event:
    current = get_event(conn, event_id)
    if current.rev != expected_rev:
        raise EditConflict(
            "This event was changed by someone else.", extra={"current_rev": current.rev}
        )
    conn.execute(
        "UPDATE events SET deleted_at = ?, rev = rev + 1, updated_at = ? WHERE id = ?",
        (_now(), _now(), event_id),
    )
    conn.commit()
    return get_event(conn, event_id, include_deleted=True)


def restore(conn: sqlite3.Connection, event_id: int) -> Event:
    ev = get_event(conn, event_id, include_deleted=True)
    m = meta(conn, event_id)
    if not m["deleted_at"]:
        return ev
    cutoff = (datetime.now(UTC) - timedelta(days=RECOVERY_DAYS)).strftime("%Y-%m-%dT%H:%M:%SZ")
    if m["deleted_at"] < cutoff:
        raise NotFound("That event was deleted more than 30 days ago.")
    conn.execute(
        "UPDATE events SET deleted_at = NULL, rev = rev + 1, updated_at = ? WHERE id = ?",
        (_now(), event_id),
    )
    conn.commit()
    return get_event(conn, event_id)


def list_deleted(conn: sqlite3.Connection) -> list[Event]:
    cutoff = (datetime.now(UTC) - timedelta(days=RECOVERY_DAYS)).strftime("%Y-%m-%dT%H:%M:%SZ")
    rows = conn.execute(
        "SELECT * FROM events WHERE calendar_id = ? AND deleted_at IS NOT NULL AND deleted_at >= ? "
        "ORDER BY deleted_at DESC",
        (CALENDAR_ID, cutoff),
    ).fetchall()
    return [row_to_event(r) for r in rows]


def purge_deleted(conn: sqlite3.Connection) -> int:
    cutoff = (datetime.now(UTC) - timedelta(days=RECOVERY_DAYS)).strftime("%Y-%m-%dT%H:%M:%SZ")
    cur = conn.execute(
        "DELETE FROM events WHERE deleted_at IS NOT NULL AND deleted_at < ?", (cutoff,)
    )
    conn.commit()
    return cur.rowcount


def add_skip(conn: sqlite3.Connection, event_id: int, day: str) -> None:
    get_event(conn, event_id)
    d = parse_date(day)
    if d is None:
        raise ValidationFailed({"date": "Choose a valid date"})
    conn.execute(
        "INSERT OR IGNORE INTO event_exceptions (event_id, occurrence_date) VALUES (?, ?)",
        (event_id, d.isoformat()),
    )
    conn.commit()


def remove_skip(conn: sqlite3.Connection, event_id: int, day: str) -> None:
    conn.execute(
        "DELETE FROM event_exceptions WHERE event_id = ? AND occurrence_date = ?", (event_id, day)
    )
    conn.commit()


def skips_for(conn: sqlite3.Connection, event_ids: list[int]) -> dict[int, set[date]]:
    if not event_ids:
        return {}
    marks = ",".join("?" * len(event_ids))
    out: dict[int, set[date]] = {}
    for r in conn.execute(
        f"SELECT event_id, occurrence_date FROM event_exceptions WHERE event_id IN ({marks})",
        event_ids,
    ):
        out.setdefault(r["event_id"], set()).add(date.fromisoformat(r["occurrence_date"]))
    return out


def list_skips(conn: sqlite3.Connection, event_id: int) -> list[date]:
    rows = conn.execute(
        "SELECT occurrence_date FROM event_exceptions WHERE event_id = ? ORDER BY occurrence_date",
        (event_id,),
    ).fetchall()
    return [date.fromisoformat(r["occurrence_date"]) for r in rows]
