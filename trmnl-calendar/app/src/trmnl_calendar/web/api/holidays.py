"""Holiday lists by region and a one-time import into a category (docs/10, slice 4)."""

from __future__ import annotations

import sqlite3
from typing import Any

from fastapi import APIRouter, Body, Depends, Request
from fastapi.responses import JSONResponse

from trmnl_calendar.domain import holidays as H
from trmnl_calendar.errors import LookupFailed, NotFound, ValidationFailed
from trmnl_calendar.holiday_source import HolidaySource, HolidaySourceError
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


def online_source(request: Request) -> HolidaySource:
    if not request.app.state.settings.holiday_lookup:
        raise NotFound("The online holiday lookup is switched off.")
    return request.app.state.holiday_source


def online_rows(source: HolidaySource, country: object, sub: object, ys: list[int]):
    if not isinstance(country, str) or len(country.strip()) != 2:
        raise ValidationFailed({"country": "Send a two-letter country code such as CA"})
    if sub is not None and not isinstance(sub, str):
        raise ValidationFailed({"subdivision": "Send a code such as CA-NS, or leave it out"})
    try:
        return source.lookup(country, (sub or "").strip(), ys)
    except HolidaySourceError as exc:
        raise LookupFailed(str(exc)) from exc


def item_json(h: H.Holiday) -> dict:
    return {
        "key": h.key,
        "name": h.name,
        "date": h.date.isoformat(),
        "group": h.group,
        "default": h.group in H.CHECKED_BY_DEFAULT,
    }


@router.get("/holidays/online/countries")
def online_countries(source: HolidaySource = Depends(online_source)) -> list[dict]:
    try:
        return [{"code": c.code, "name": c.name} for c in source.countries()]
    except HolidaySourceError as exc:
        raise LookupFailed(str(exc)) from exc


@router.get("/holidays/online")
def online_list(
    country: str,
    years: str,
    subdivision: str = "",
    source: HolidaySource = Depends(online_source),
) -> dict:
    ys = parse_years(years)
    codes, rows = online_rows(source, country, subdivision, ys)
    return {
        "country": country.strip().upper(),
        "subdivision": subdivision.strip().upper(),
        "subdivisions": codes,
        "groups": [{"id": g, "label": n} for g, n in H.GROUPS],
        "items": [item_json(h) for h in rows],
    }


@router.get("/holidays/regions")
def regions() -> list[dict]:
    return [{"id": k, "label": v["label"], "note": v["note"]} for k, v in H.REGIONS.items()]


@router.get("/holidays")
def list_holidays(region: str, years: str, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    region_id, ys = parse_region(region), parse_years(years)
    out = [item_json(h) for y in ys for h in H.holidays_for(region_id, y)]
    return {
        "region": region_id,
        "groups": [{"id": g, "label": n} for g, n in H.GROUPS],
        "items": out,
    }


@router.post("/holidays/import")
def import_holidays(
    request: Request,
    body: dict[str, Any] = Body(...),
    conn: sqlite3.Connection = Depends(get_conn),
) -> JSONResponse:
    unknown = set(body) - {"region", "country", "subdivision", "years", "keys", "category_id"}
    if unknown:
        raise ValidationFailed({k: "Unknown field" for k in sorted(unknown)})
    online = body.get("country") is not None
    if online and body.get("region") is not None:
        raise ValidationFailed({"region": "Send a region or a country, not both"})
    ys = parse_years(body.get("years"))
    region_id = "" if online else parse_region(body.get("region"))
    keys = body.get("keys")
    if keys is not None and (
        not isinstance(keys, list) or not all(isinstance(k, str) for k in keys)
    ):
        raise ValidationFailed({"keys": "Send a list of holiday keys, or leave it out"})
    cid = body.get("category_id")
    if cid is not None and (isinstance(cid, bool) or not isinstance(cid, int)):
        raise ValidationFailed({"category_id": "Send a category id, or leave it out"})
    if online:
        _, pool = online_rows(
            online_source(request), body.get("country"), body.get("subdivision"), ys
        )
    else:
        pool = [h for y in ys for h in H.holidays_for(region_id, y)]
    items = [
        h for h in pool if (h.key in keys if keys is not None else h.group in H.CHECKED_BY_DEFAULT)
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
