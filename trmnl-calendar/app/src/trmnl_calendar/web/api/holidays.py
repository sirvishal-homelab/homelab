"""Holiday lists by region and a one-time import into a category (docs/10, slice 4)."""

from __future__ import annotations

import sqlite3
from typing import Any

from fastapi import APIRouter, Body, Depends
from fastapi.responses import JSONResponse

from trmnl_calendar.domain import holidays as H
from trmnl_calendar.errors import ValidationFailed
from trmnl_calendar.storage.repositories import holidays as repo
from trmnl_calendar.web.deps import get_conn

router = APIRouter(prefix="/api/v1", tags=["holidays"])
MAX_YEARS = 5


def parse_years(raw: object) -> list[int]:
    if isinstance(raw, str):
        raw = [p for p in raw.replace(" ", "").split(",") if p]
    if not isinstance(raw, list) or not raw:
        raise ValidationFailed({"years": "Choose at least one year"})
    try:
        years = sorted({int(y) for y in raw})
    except (TypeError, ValueError):
        raise ValidationFailed({"years": "Years must be numbers like 2026"}) from None
    if len(years) > MAX_YEARS or any(not 1971 <= y <= 2099 for y in years):
        raise ValidationFailed({"years": f"Choose up to {MAX_YEARS} years between 1971 and 2099"})
    return years


def parse_region(raw: object) -> str:
    if raw not in H.REGIONS:
        raise ValidationFailed({"region": "Choose a region from the list"})
    return str(raw)


@router.get("/holidays/regions")
def regions() -> list[dict]:
    return [{"id": k, "label": v["label"], "note": v["note"]} for k, v in H.REGIONS.items()]


@router.get("/holidays")
def list_holidays(region: str, years: str, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    region_id, ys = parse_region(region), parse_years(years)
    out = [
        {
            "key": h.key,
            "name": h.name,
            "date": h.date.isoformat(),
            "group": h.group,
            "default": h.group in H.CHECKED_BY_DEFAULT,
        }
        for y in ys
        for h in H.holidays_for(region_id, y)
    ]
    return {
        "region": region_id,
        "groups": [{"id": g, "label": n} for g, n in H.GROUPS],
        "items": out,
    }


@router.post("/holidays/import")
def import_holidays(
    body: dict[str, Any] = Body(...),
    conn: sqlite3.Connection = Depends(get_conn),
) -> JSONResponse:
    unknown = set(body) - {"region", "years", "keys", "category_id"}
    if unknown:
        raise ValidationFailed({k: "Unknown field" for k in sorted(unknown)})
    region_id, ys = parse_region(body.get("region")), parse_years(body.get("years"))
    keys = body.get("keys")
    if keys is not None and (
        not isinstance(keys, list) or not all(isinstance(k, str) for k in keys)
    ):
        raise ValidationFailed({"keys": "Send a list of holiday keys, or leave it out"})
    cid = body.get("category_id")
    if cid is not None and (isinstance(cid, bool) or not isinstance(cid, int)):
        raise ValidationFailed({"category_id": "Send a category id, or leave it out"})
    items = [
        h
        for y in ys
        for h in H.holidays_for(region_id, y)
        if (h.key in keys if keys is not None else h.group in H.CHECKED_BY_DEFAULT)
    ]
    res = repo.import_holidays(conn, items, cid)
    return JSONResponse(
        {
            "added": res.added,
            "skipped": res.skipped,
            "category": {"id": res.category_id, "name": res.category_name},
        },
        status_code=201 if res.added else 200,
    )
