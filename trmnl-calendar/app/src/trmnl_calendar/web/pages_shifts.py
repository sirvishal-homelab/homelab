"""Shifts pages: an overview of rotating schedules, and a work month (docs/20, ADR-0014)."""

from __future__ import annotations

from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse

from trmnl_calendar.web import shifts_view, views
from trmnl_calendar.web.deps import chrome, get_conn, now_utc, who_from_cookie

router = APIRouter(include_in_schema=False)


def _page(request: Request, name: str, **ctx) -> HTMLResponse:
    return HTMLResponse(views.render_page(name, **{**chrome(request), "nav": "shifts", **ctx}))


@router.get("/shifts")
def shifts(request: Request, saved: str = "", conn=Depends(get_conn)):
    ctx = shifts_view.overview_context(conn, now_utc(request), saved)
    return _page(request, "shifts.html", **ctx)


@router.get("/shifts/month")
def shifts_month(
    request: Request,
    e: int | None = None,
    y: int | None = None,
    m: int | None = None,
    day: str | None = None,
    saved: str = "",
    conn=Depends(get_conn),
):
    me = who_from_cookie(request, conn)
    ctx = shifts_view.month_context(
        conn, now_utc(request), e, y, m, day, me.id if me else None, saved
    )
    return _page(request, "shifts_month.html", **ctx)
