"""Categories (sections): create, edit, order, delete with a move target (docs/03, ADR-0004)."""

from __future__ import annotations

import sqlite3

from trmnl_calendar.domain.models import Section
from trmnl_calendar.errors import InUse, NotFound, ValidationFailed
from trmnl_calendar.storage.repositories.calendar import CALENDAR_ID, list_sections

MAX_SECTIONS = 8
PLACEMENTS = ("lane", "banner")


def get_section(conn: sqlite3.Connection, section_id: int) -> Section:
    for s in list_sections(conn):
        if s.id == section_id:
            return s
    raise NotFound("That category no longer exists.")


def count_events(conn: sqlite3.Connection, section_id: int) -> int:
    return conn.execute(
        "SELECT COUNT(*) FROM events WHERE section_id = ?", (section_id,)
    ).fetchone()[0]


def _name_errors(conn: sqlite3.Connection, name: object, own_id: int | None) -> tuple[str, dict]:
    n = name.strip() if isinstance(name, str) else ""
    if not 1 <= len(n) <= 30:
        return n, {"name": "Enter a name up to 30 characters"}
    clash = conn.execute(
        "SELECT id FROM sections WHERE calendar_id = ? AND name = ? COLLATE NOCASE AND id IS NOT ?",
        (CALENDAR_ID, n, own_id),
    ).fetchone()
    return n, ({"name": "You already have a category with that name"} if clash else {})


def _weight(value: object) -> tuple[int, dict]:
    try:
        w = int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return 1, {"weight": "Choose a height from 1 to 4"}
    return w, ({} if 1 <= w <= 4 else {"weight": "Choose a height from 1 to 4"})


def create_section(
    conn: sqlite3.Connection,
    name: object,
    *,
    weight: object = 1,
    visible: bool = True,
    placement: str = "lane",
) -> Section:
    n, err = _name_errors(conn, name, None)
    w, werr = _weight(weight)
    err.update(werr)
    if placement not in PLACEMENTS:
        err["placement"] = "Choose lane or banner"
    if len(list_sections(conn)) >= MAX_SECTIONS:
        err["name"] = f"You can have up to {MAX_SECTIONS} categories"
    if err:
        raise ValidationFailed(err)
    pos = conn.execute(
        "SELECT COALESCE(MAX(position), 0) + 1 FROM sections WHERE calendar_id = ?", (CALENDAR_ID,)
    ).fetchone()[0]
    cur = conn.execute(
        "INSERT INTO sections (calendar_id, name, position, weight, visible, placement) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (CALENDAR_ID, n, pos, w, int(bool(visible)), placement),
    )
    conn.commit()
    return get_section(conn, cur.lastrowid)


def update_section(conn: sqlite3.Connection, section_id: int, patch: dict) -> Section:
    cur = get_section(conn, section_id)
    unknown = set(patch) - {"name", "weight", "visible", "placement"}
    if unknown:
        raise ValidationFailed({k: "Unknown field" for k in sorted(unknown)})
    n, err = _name_errors(conn, patch.get("name", cur.name), section_id)
    w, werr = _weight(patch.get("weight", cur.weight))
    err.update(werr)
    placement = patch.get("placement", cur.placement)
    if placement not in PLACEMENTS:
        err["placement"] = "Choose lane or banner"
    visible = patch.get("visible", cur.visible)
    if not isinstance(visible, bool):
        err["visible"] = "Must be true or false"
    if err:
        raise ValidationFailed(err)
    conn.execute(
        "UPDATE sections SET name = ?, weight = ?, visible = ?, placement = ? WHERE id = ?",
        (n, w, int(visible), placement, section_id),
    )
    conn.commit()
    return get_section(conn, section_id)


def delete_section(conn: sqlite3.Connection, section_id: int, move_to: int | None = None) -> int:
    """Delete a category; its events move to move_to (required if any). Returns events moved."""
    get_section(conn, section_id)
    if len(list_sections(conn)) <= 1:
        raise InUse("You need at least one category.")
    count = conn.execute(
        "SELECT COUNT(*) FROM events WHERE section_id = ?", (section_id,)
    ).fetchone()[0]
    if count:
        if move_to is None or move_to == section_id:
            raise InUse(f"{count} event(s) use this category. Choose where they should move.")
        get_section(conn, move_to)
        conn.execute("UPDATE events SET section_id = ? WHERE section_id = ?", (move_to, section_id))
    conn.execute("DELETE FROM sections WHERE id = ?", (section_id,))
    conn.commit()
    return count


def reorder_sections(conn: sqlite3.Connection, ids: list[int]) -> None:
    existing = [s.id for s in list_sections(conn)]
    if sorted(ids) != sorted(existing) or len(set(ids)) != len(ids):
        raise ValidationFailed({"ids": "Send every category id exactly once"})
    for pos, sid in enumerate(ids, start=1):
        conn.execute("UPDATE sections SET position = ? WHERE id = ?", (pos, sid))
    conn.commit()


def move_section(conn: sqlite3.Connection, section_id: int, direction: int) -> None:
    """Swap with the neighbour above (-1) or below (+1); no-op at the ends."""
    ids = [s.id for s in list_sections(conn)]
    if section_id not in ids:
        raise NotFound("That category no longer exists.")
    i = ids.index(section_id)
    j = i + direction
    if 0 <= j < len(ids):
        ids[i], ids[j] = ids[j], ids[i]
        reorder_sections(conn, ids)
