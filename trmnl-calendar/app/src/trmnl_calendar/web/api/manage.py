"""Categories, people, calendar settings and "who am I" (docs/06 section 2)."""

from __future__ import annotations

import sqlite3
from typing import Any

from fastapi import APIRouter, Body, Depends, Request
from fastapi.responses import JSONResponse

from trmnl_calendar.domain.models import CalendarInfo, Person, Section
from trmnl_calendar.errors import ValidationFailed
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import people as ppl
from trmnl_calendar.storage.repositories import sections as sec
from trmnl_calendar.web.deps import WHO_COOKIE, get_conn, who_from_cookie

router = APIRouter(prefix="/api/v1", tags=["manage"])


def _cat(s: Section) -> dict:
    return {
        "id": s.id,
        "name": s.name,
        "position": s.position,
        "weight": s.weight,
        "visible": s.visible,
        "placement": s.placement,
    }


def _person(p: Person) -> dict:
    return {
        "id": p.id,
        "name": p.name,
        "label": p.label,
        "position": p.position,
        "active": p.active,
    }


def _calendar(c: CalendarInfo) -> dict:
    return {
        "name": c.name,
        "timezone": c.timezone,
        "week_start": c.week_start,
        "time_format": c.time_format,
    }


# ---- categories


@router.get("/categories")
def list_categories(conn: sqlite3.Connection = Depends(get_conn)) -> list[dict]:
    return [_cat(s) for s in cal.list_sections(conn)]


@router.post("/categories", status_code=201)
def create_category(
    body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> JSONResponse:
    s = sec.create_section(
        conn,
        body.get("name"),
        weight=body.get("weight", 1),
        visible=body.get("visible", True),
        placement=body.get("placement", "lane"),
    )
    return JSONResponse(_cat(s), status_code=201)


@router.put("/categories/order")
def order_categories(
    body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> list[dict]:
    ids = body.get("ids")
    if not isinstance(ids, list) or not all(isinstance(i, int) for i in ids):
        raise ValidationFailed({"ids": "Send a list of category ids"})
    sec.reorder_sections(conn, ids)
    return [_cat(s) for s in cal.list_sections(conn)]


@router.patch("/categories/{section_id}")
def patch_category(
    section_id: int, body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> dict:
    return _cat(sec.update_section(conn, section_id, body))


@router.delete("/categories/{section_id}")
def delete_category(
    section_id: int, move_to: int | None = None, conn: sqlite3.Connection = Depends(get_conn)
) -> dict:
    moved = sec.delete_section(conn, section_id, move_to)
    return {"deleted": section_id, "events_moved": moved}


# ---- people


@router.get("/people")
def list_people(conn: sqlite3.Connection = Depends(get_conn)) -> list[dict]:
    return [_person(p) for p in ppl.list_people(conn)]


@router.post("/people", status_code=201)
def create_person(
    body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> JSONResponse:
    return JSONResponse(
        _person(ppl.create_person(conn, body.get("name"), body.get("label"))), status_code=201
    )


@router.patch("/people/{person_id}")
def patch_person(
    person_id: int, body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> dict:
    return _person(ppl.update_person(conn, person_id, body))


@router.delete("/people/{person_id}")
def delete_person(person_id: int, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    ppl.delete_person(conn, person_id)
    return {"deleted": person_id}


# ---- calendar and me


@router.get("/calendar")
def get_calendar(conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    return _calendar(cal.get_calendar(conn))


@router.patch("/calendar")
def patch_calendar(
    body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> dict:
    return _calendar(cal.update_calendar(conn, body))


@router.get("/me")
def get_me(request: Request, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    me = who_from_cookie(request, conn)
    return {"person": _person(me) if me else None}


@router.put("/me")
def put_me(
    body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> JSONResponse:
    pid = body.get("person_id")
    person = None
    if pid is not None:
        if isinstance(pid, bool) or not isinstance(pid, int):
            raise ValidationFailed({"person_id": "Send a person id or null"})
        person = ppl.get_person(conn, pid)
    resp = JSONResponse({"person": _person(person) if person else None})
    resp.set_cookie(
        WHO_COOKIE,
        str(person.id if person else 0),
        max_age=365 * 24 * 3600,
        samesite="lax",
        httponly=True,
    )
    return resp
