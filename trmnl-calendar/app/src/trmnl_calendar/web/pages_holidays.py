"""The holiday picker: choose a region and years, tick the holidays, add them once."""

from __future__ import annotations

from contextlib import closing
from urllib.parse import quote

from fastapi import APIRouter, Depends, Query, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from starlette.concurrency import run_in_threadpool

from trmnl_calendar.domain import holidays as H
from trmnl_calendar.errors import NotFound, ValidationFailed
from trmnl_calendar.holiday_source import HolidaySource, HolidaySourceError
from trmnl_calendar.storage import db
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import holidays as repo
from trmnl_calendar.web import forms, render_service, views
from trmnl_calendar.web.deps import chrome, get_conn, now_utc

router = APIRouter(include_in_schema=False)


def _page(request: Request, status: int = 200, **ctx) -> HTMLResponse:
    return HTMLResponse(
        views.render_page("holidays.html", **{**chrome(request), "nav": "categories", **ctx}),
        status_code=status,
    )


def _this_year(request: Request, conn) -> int:
    tz = cal.get_calendar(conn).timezone
    return render_service.today_for(tz, now_utc(request)).year


def _years(values: list[str], allowed: list[int]) -> list[int]:
    return sorted({int(v) for v in values if v.isdigit() and int(v) in allowed})


@router.get("/holidays")
def holidays(
    request: Request,
    region: str = "",
    year: list[str] = Query(default=[]),
    done: str = "",
    conn=Depends(get_conn),
):
    this_year = _this_year(request, conn)
    choices = [this_year, this_year + 1, this_year + 2]
    ys = _years(list(year), choices) or ([this_year] if region else [])
    ctx = views.holidays_context(
        conn,
        region,
        ys,
        this_year,
        done=done[:120],
        online_enabled=request.app.state.settings.holiday_lookup,
    )
    return _page(request, **ctx)


def _source(request: Request) -> HolidaySource:
    if not request.app.state.settings.holiday_lookup:
        raise NotFound("The online holiday lookup is switched off.")
    return request.app.state.holiday_source


def _sub_choices(codes: list[str]) -> list[tuple[str, str]]:
    return [(c, H.REGIONS[c]["label"] if c in H.REGIONS else c) for c in codes]


def _online_ctx(request, conn, source, country, sub, ys, this_year, **extra) -> dict:
    """Everything the lookup page needs; a failed lookup becomes a message, not an error page."""
    error = extra.pop("error_message", "")
    countries, codes, items = [], [], []
    try:
        countries = source.countries()
    except HolidaySourceError as exc:
        error = error or str(exc)
    if country and ys and not error:
        try:
            codes, items = source.lookup(country, sub, ys)
        except HolidaySourceError as exc:
            error = str(exc)
    return views.holidays_online_context(
        conn,
        countries=countries,
        country=country,
        subdivision=sub,
        subdivisions=_sub_choices(codes),
        items=items,
        years=ys,
        this_year=this_year,
        error_message=error,
        **extra,
    )


@router.get("/holidays/online")
def holidays_online(
    request: Request,
    country: str = "",
    subdivision: str = "",
    year: list[str] = Query(default=[]),
    done: str = "",
    conn=Depends(get_conn),
):
    source = _source(request)
    this_year = _this_year(request, conn)
    ys = _years(list(year), [this_year, this_year + 1, this_year + 2])
    country = country.strip().upper()[:2]
    ys = ys or ([this_year] if country else [])
    ctx = _online_ctx(
        request,
        conn,
        source,
        country,
        subdivision.strip().upper()[:12],
        ys,
        this_year,
        done=done[:120],
    )
    return _page_online(request, **ctx)


def _page_online(request: Request, status: int = 200, **ctx) -> HTMLResponse:
    return HTMLResponse(
        views.render_page(
            "holidays_online.html", **{**chrome(request), "nav": "categories", **ctx}
        ),
        status_code=status,
    )


@router.post("/holidays/import")
async def holidays_import(request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with closing(db.connect(request.app.state.settings.db_path)) as conn:
            this_year = _this_year(request, conn)
            region = (form.get("region") or [""])[0]
            cat = (form.get("category_id") or [""])[0]
            category_id = int(cat) if cat.isdigit() else None
            tokens = [t for t in form.get("h", []) if ":" in t]
            ys = sorted({int(t.split(":")[1]) for t in tokens if t.split(":")[1].isdigit()})
            country = (form.get("country") or [""])[0].strip().upper()
            sub = (form.get("subdivision") or [""])[0].strip().upper()
            try:
                wanted = set(tokens)
                if country:
                    source = _source(request)
                    try:
                        _, rows = source.lookup(country, sub, ys) if ys else ([], [])
                    except HolidaySourceError as exc:
                        raise ValidationFailed({"country": str(exc)}) from exc
                    items = [h for h in rows if f"{h.key}:{h.date.year}" in wanted]
                else:
                    if region not in H.REGIONS:
                        raise ValidationFailed({"region": "Choose a region"})
                    items = [
                        h
                        for y in ys
                        for h in H.holidays_for(region, y)
                        if f"{h.key}:{h.date.year}" in wanted
                    ]
                res = repo.import_holidays(conn, items, category_id)
            except ValidationFailed as exc:
                if country:
                    ctx = _online_ctx(
                        request,
                        conn,
                        request.app.state.holiday_source,
                        country,
                        sub,
                        ys or [this_year],
                        this_year,
                        category_id=category_id,
                        error_message=" ".join(exc.fields.values()),
                    )
                    return _page_online(request, 422, **ctx)
                ctx = views.holidays_context(
                    conn,
                    region,
                    ys or [this_year],
                    this_year,
                    category_id,
                    error_message=" ".join(exc.fields.values()),
                )
                return _page(request, 422, **ctx)
            msg = f"Added {res.added} to “{res.category_name}”"
            if res.skipped:
                msg += f"; {res.skipped} already there"
            target = "/holidays/online" if country else "/holidays"
            return RedirectResponse(target + "?done=" + quote(msg + "."), status_code=303)

    return await run_in_threadpool(work)
