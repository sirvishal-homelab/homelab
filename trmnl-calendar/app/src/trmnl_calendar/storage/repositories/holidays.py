"""Add holidays from a region list as ordinary events in one category (no duplicates)."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass

from trmnl_calendar.domain.holidays import Holiday
from trmnl_calendar.errors import ValidationFailed
from trmnl_calendar.storage.repositories import events as events_repo
from trmnl_calendar.storage.repositories import sections as sec
from trmnl_calendar.storage.repositories.calendar import CALENDAR_ID, list_sections

DEFAULT_CATEGORY = "Holidays"


@dataclass(frozen=True)
class ImportResult:
    added: int
    skipped: int
    category_id: int
    category_name: str


def target_category(conn: sqlite3.Connection, category_id: int | None) -> tuple[int, str]:
    """The chosen category, else 'Holidays' (created as a banner if it does not exist)."""
    sections = list_sections(conn)
    if category_id is not None:
        for s in sections:
            if s.id == category_id:
                return s.id, s.name
        raise ValidationFailed({"category_id": "That category no longer exists"})
    for s in sections:
        if s.name.lower() == DEFAULT_CATEGORY.lower():
            return s.id, s.name
    made = sec.create_section(conn, DEFAULT_CATEGORY, placement="banner")
    return made.id, made.name


def import_holidays(
    conn: sqlite3.Connection, items: list[Holiday], category_id: int | None = None
) -> ImportResult:
    if not items:
        raise ValidationFailed({"holidays": "Tick at least one holiday"})
    cid, cname = target_category(conn, category_id)
    added = skipped = 0
    for h in items:
        exists = conn.execute(
            "SELECT 1 FROM events WHERE calendar_id = ? AND section_id = ? AND deleted_at IS NULL "
            "AND start_date = ? AND title = ? COLLATE NOCASE",
            (CALENDAR_ID, cid, h.date.isoformat(), h.name),
        ).fetchone()
        if exists:
            skipped += 1
            continue
        events_repo.create_event(
            conn, {"title": h.name, "section_id": cid, "start_date": h.date.isoformat()}
        )
        added += 1
    return ImportResult(added, skipped, cid, cname)
