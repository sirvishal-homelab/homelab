"""People: the household members an event can be for (docs/03)."""

from __future__ import annotations

import sqlite3

from trmnl_calendar.domain.models import Person
from trmnl_calendar.errors import InUse, NotFound, ValidationFailed
from trmnl_calendar.storage.repositories.calendar import CALENDAR_ID

MAX_PEOPLE = 10


def _row(r: sqlite3.Row) -> Person:
    return Person(r["id"], r["name"], r["label"], r["position"], bool(r["active"]))


def list_people(conn: sqlite3.Connection, *, include_inactive: bool = True) -> list[Person]:
    sql = "SELECT * FROM people WHERE calendar_id = ?"
    if not include_inactive:
        sql += " AND active = 1"
    return [_row(r) for r in conn.execute(sql + " ORDER BY position, id", (CALENDAR_ID,))]


def get_person(conn: sqlite3.Connection, person_id: int) -> Person:
    r = conn.execute(
        "SELECT * FROM people WHERE id = ? AND calendar_id = ?", (person_id, CALENDAR_ID)
    ).fetchone()
    if not r:
        raise NotFound("That person no longer exists.")
    return _row(r)


def _clean(name: object, label: object, conn: sqlite3.Connection, own_id: int | None):
    err: dict[str, str] = {}
    n = name.strip() if isinstance(name, str) else ""
    lab = label.strip() if isinstance(label, str) else ""
    if not 1 <= len(n) <= 30:
        err["name"] = "Enter a name up to 30 characters"
    if not 1 <= len(lab) <= 2:
        err["label"] = "Use one or two letters, for example A"
    else:
        clash = conn.execute(
            "SELECT id FROM people "
            "WHERE calendar_id = ? AND label = ? COLLATE NOCASE AND id IS NOT ?",
            (CALENDAR_ID, lab, own_id),
        ).fetchone()
        if clash:
            err["label"] = "Another person already uses that label"
    return n, lab, err


def create_person(conn: sqlite3.Connection, name: object, label: object) -> Person:
    n, lab, err = _clean(name, label, conn, None)
    count = conn.execute(
        "SELECT COUNT(*) FROM people WHERE calendar_id = ?", (CALENDAR_ID,)
    ).fetchone()[0]
    if count >= MAX_PEOPLE:
        err["name"] = f"You can have up to {MAX_PEOPLE} people"
    if err:
        raise ValidationFailed(err)
    pos = conn.execute(
        "SELECT COALESCE(MAX(position), 0) + 1 FROM people WHERE calendar_id = ?", (CALENDAR_ID,)
    ).fetchone()[0]
    cur = conn.execute(
        "INSERT INTO people (calendar_id, name, label, position) VALUES (?, ?, ?, ?)",
        (CALENDAR_ID, n, lab, pos),
    )
    conn.commit()
    return get_person(conn, cur.lastrowid)


def update_person(conn: sqlite3.Connection, person_id: int, patch: dict) -> Person:
    cur = get_person(conn, person_id)
    unknown = set(patch) - {"name", "label", "active"}
    if unknown:
        raise ValidationFailed({k: "Unknown field" for k in sorted(unknown)})
    n, lab, err = _clean(
        patch.get("name", cur.name), patch.get("label", cur.label), conn, person_id
    )
    active = patch.get("active", cur.active)
    if not isinstance(active, bool):
        err["active"] = "Must be true or false"
    if err:
        raise ValidationFailed(err)
    conn.execute(
        "UPDATE people SET name = ?, label = ?, active = ? WHERE id = ?",
        (n, lab, int(active), person_id),
    )
    conn.commit()
    return get_person(conn, person_id)


def delete_person(conn: sqlite3.Connection, person_id: int) -> None:
    get_person(conn, person_id)
    used = conn.execute(
        "SELECT COUNT(DISTINCT event_id) FROM event_people WHERE person_id = ?", (person_id,)
    ).fetchone()[0]
    if used:
        raise InUse(
            f"{used} event(s) use this person. Mark them as not active instead; "
            "old events stay intact."
        )
    conn.execute("DELETE FROM people WHERE id = ?", (person_id,))
    conn.commit()
