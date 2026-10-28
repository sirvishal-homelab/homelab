"""Events, occurrences and categories API (docs/06). PATCH not PUT, If-Match, Idempotency-Key."""

from __future__ import annotations

import sqlite3
from datetime import date, timedelta
from typing import Any

from fastapi import APIRouter, Body, Depends, Header, Query, Request
from fastapi.responses import JSONResponse

from trmnl_calendar.domain.recurrence import MAX_WINDOW_DAYS, occurrences
from trmnl_calendar.domain.validation import parse_date
from trmnl_calendar.errors import AppError
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import idempotency
from trmnl_calendar.storage.repositories import people as people_repo
from trmnl_calendar.web.deps import get_conn, parse_if_match, who_from_cookie
from trmnl_calendar.web.serialize import event_to_api, patch_from_api

router = APIRouter(prefix="/api/v1", tags=["events"])


def _out(conn: sqlite3.Connection, ev) -> dict:
    people = {p.id: p for p in people_repo.list_people(conn)}
    return event_to_api(ev, repo.meta(conn, ev.id), people)


def _me_id(request: Request, conn: sqlite3.Connection) -> int | None:
    me = who_from_cookie(request, conn)
    return me.id if me else None


def _window(start: str | None, end: str | None) -> tuple[date, date]:
    a, b = parse_date(start), parse_date(end)
    if a is None or b is None:
        raise AppError("Give both 'from' and 'to' as YYYY-MM-DD.")
    if b < a or (b - a).days > MAX_WINDOW_DAYS:
        raise AppError(f"The window must be 0 to {MAX_WINDOW_DAYS} days.")
    return a, b


@router.get("/events")
def list_events(
    from_: str | None = Query(None, alias="from"),
    to: str | None = None,
    deleted: bool = False,
    conn: sqlite3.Connection = Depends(get_conn),
) -> list[dict]:
    if deleted:
        evs = repo.list_deleted(conn)
    else:
        evs = repo.list_events(conn, *_window(from_, to))
    return [_out(conn, e) for e in evs]


@router.post("/events", status_code=201)
def create_event(
    request: Request,
    body: dict[str, Any] = Body(...),
    idempotency_key: str | None = Header(None, alias="Idempotency-Key"),
    conn: sqlite3.Connection = Depends(get_conn),
) -> JSONResponse:
    scope = "POST /api/v1/events"
    if idempotency_key:
        hit = idempotency.lookup(conn, idempotency_key, scope)
        if hit:
            return JSONResponse(hit[1], status_code=hit[0])
    ev = repo.create_event(conn, patch_from_api(body), person_id=_me_id(request, conn))
    out = _out(conn, ev)
    if idempotency_key:
        idempotency.store(conn, idempotency_key, scope, 201, out)
    return JSONResponse(out, status_code=201)


@router.get("/events/{event_id}")
def get_event(event_id: int, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    ev = repo.get_event(conn, event_id)
    return _out(conn, ev)


@router.patch("/events/{event_id}")
def patch_event(
    event_id: int,
    request: Request,
    body: dict[str, Any] = Body(...),
    if_match: str | None = Header(None),
    conn: sqlite3.Connection = Depends(get_conn),
) -> dict:
    rev = parse_if_match(if_match)
    ev = repo.update_event(
        conn, event_id, patch_from_api(body), rev, person_id=_me_id(request, conn)
    )
    return _out(conn, ev)


@router.delete("/events/{event_id}")
def delete_event(
    event_id: int,
    if_match: str | None = Header(None),
    conn: sqlite3.Connection = Depends(get_conn),
) -> dict:
    ev = repo.soft_delete(conn, event_id, parse_if_match(if_match))
    return _out(conn, ev)


@router.post("/events/{event_id}/restore")
def restore_event(event_id: int, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    ev = repo.restore(conn, event_id)
    return _out(conn, ev)


@router.post("/events/{event_id}/skips", status_code=201)
def add_skip(
    event_id: int, body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> dict:
    repo.add_skip(conn, event_id, str(body.get("date", "")))
    return {
        "event_id": event_id,
        "skipped": [d.isoformat() for d in repo.list_skips(conn, event_id)],
    }


@router.delete("/events/{event_id}/skips/{day}")
def remove_skip(event_id: int, day: str, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    repo.get_event(conn, event_id)
    repo.remove_skip(conn, event_id, day)
    return {
        "event_id": event_id,
        "skipped": [d.isoformat() for d in repo.list_skips(conn, event_id)],
    }


@router.get("/occurrences")
def list_occurrences(
    from_: str | None = Query(None, alias="from"),
    to: str | None = None,
    conn: sqlite3.Connection = Depends(get_conn),
) -> list[dict]:
    a, b = _window(from_, to)
    calendar = cal.get_calendar(conn)
    evs = repo.list_events(conn, a - timedelta(days=0), b)
    occ = occurrences(evs, repo.skips_for(conn, [e.id for e in evs]), a, b, calendar.week_start)
    return [
        {
            "event_id": o.event_id,
            "title": o.title,
            "category_id": o.section_id,
            "date": o.date.isoformat(),
            "start_time": o.start_time,
            "end_time": o.end_time,
            "all_day": o.all_day,
            "location": o.location,
            "day_number": o.day_number,
            "day_count": o.day_count,
        }
        for o in occ
    ]
