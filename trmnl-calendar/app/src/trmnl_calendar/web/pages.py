"""Server-rendered pages and the image routes (docs/06 section 4)."""

from __future__ import annotations

import re
from contextlib import closing
from urllib.parse import quote

from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from starlette.concurrency import run_in_threadpool

from trmnl_calendar.errors import EditConflict, NotFound, ValidationFailed
from trmnl_calendar.storage import db
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import views as views_repo
from trmnl_calendar.web import forms, render_service, views
from trmnl_calendar.web.deps import (
    chrome,
    get_conn,
    now_utc,
    safe_next,
    who_chosen,
    who_from_cookie,
)

router = APIRouter(include_in_schema=False)
_HASH = re.compile(r"^[0-9a-f]{24}\.png$")


def _html(request: Request, name: str, status: int = 200, **ctx) -> HTMLResponse:
    page = views.render_page(name, **{**chrome(request), "nav": "home", **ctx})
    return HTMLResponse(page, status_code=status)


def _see_other(url: str) -> RedirectResponse:
    return RedirectResponse(url, status_code=303)


def _open(request: Request):
    return closing(db.connect(request.app.state.settings.db_path))


def _quote(text: str) -> str:
    return quote(text)


def _back(form: dict) -> str:
    """Where to return after saving: a same-site path chosen by the page that opened the form."""
    raw = (form.get("back") or [""])[0]
    return safe_next(raw) if raw else ""


def _return_to(back: str, **params: str) -> str:
    """The home page (as before), or the page the form was opened from, with a flash message."""
    if back and back != "/":
        return views.add_query(back, **params)
    return "/?" + "&".join(f"{k}={_quote(v)}" for k, v in params.items())


def _rev(form: dict) -> int:
    try:
        return int((form.get("rev") or [""])[0])
    except ValueError:
        return -1


@router.get("/")
def home(
    request: Request,
    deleted: int | None = None,
    saved: str | None = None,
    person: int | None = None,
    conn=Depends(get_conn),
):
    flash_deleted = None
    if deleted is not None:
        try:
            ev = repo.get_event(conn, deleted, include_deleted=True)
            flash_deleted = {"id": ev.id, "title": ev.title}
        except NotFound:
            pass
    model = views.upcoming_groups(conn, now_utc(request), person)
    return _html(
        request,
        "home.html",
        flash=(saved or "")[:80],
        flash_deleted=flash_deleted,
        show_who_prompt=bool(model["people_list"]) and not who_chosen(request),
        active_view=views_repo.active_view(conn),
        **model,
    )


@router.get("/month")
def month(
    request: Request,
    y: int | None = None,
    m: int | None = None,
    day: str | None = None,
    person: int | None = None,
    saved: str | None = None,
    deleted: int | None = None,
    conn=Depends(get_conn),
):
    flash_deleted = None
    if deleted is not None:
        try:
            ev = repo.get_event(conn, deleted, include_deleted=True)
            flash_deleted = {"id": ev.id, "title": ev.title}
        except NotFound:
            pass
    ctx = views.month_context(conn, now_utc(request), y, m, person, day)
    here = f"/month?y={ctx['year']}&m={ctx['month']}" + (f"&person={person}" if person else "")
    return _html(
        request,
        "month.html",
        flash=(saved or "")[:80],
        flash_deleted=flash_deleted,
        back=here,
        **ctx,
    )


@router.get("/events/new")
def new_form(
    request: Request, date: str | None = None, back: str | None = None, conn=Depends(get_conn)
):
    calendar = cal.get_calendar(conn)
    today = render_service.today_for(calendar.timezone, now_utc(request))
    sections = cal.list_sections(conn, include_hidden=False)
    me = who_from_cookie(request, conn)
    values = forms.event_to_values(
        None,
        date or today.isoformat(),
        sections[0].id if sections else None,
        (me.id,) if me else (),
    )
    ctx = views.form_context(conn, values)
    return _html(request, "event_form.html", back=safe_next(back) if back else "", **ctx)


@router.post("/events/new")
async def create(request: Request):
    form = forms.parse_body(await request.body())
    back = _back(form)

    def work() -> Response:
        with _open(request) as conn:
            me = who_from_cookie(request, conn)
            try:
                ev = repo.create_event(
                    conn, forms.form_to_event_data(form), person_id=me.id if me else None
                )
            except ValidationFailed as exc:
                ctx = views.form_context(conn, forms.values_from_form(form), exc.fields)
                return _html(request, "event_form.html", 422, back=back, **ctx)
            return _see_other(_return_to(back, saved=f"Added “{ev.title}”."))

    return await run_in_threadpool(work)


@router.get("/events/{event_id}/edit")
def edit_form(event_id: int, request: Request, back: str | None = None, conn=Depends(get_conn)):
    ev = repo.get_event(conn, event_id)
    ctx = views.form_context(
        conn, forms.event_to_values(ev), event=ev, skips=repo.list_skips(conn, ev.id)
    )
    return _html(request, "event_form.html", back=safe_next(back) if back else "", **ctx)


@router.post("/events/{event_id}/edit")
async def edit(event_id: int, request: Request):
    form = forms.parse_body(await request.body())
    back = _back(form)

    def work() -> Response:
        with _open(request) as conn:
            me = who_from_cookie(request, conn)
            try:
                ev = repo.update_event(
                    conn,
                    event_id,
                    forms.form_to_event_data(form),
                    _rev(form),
                    person_id=me.id if me else None,
                )
            except ValidationFailed as exc:
                ev = repo.get_event(conn, event_id)
                ctx = views.form_context(
                    conn,
                    forms.values_from_form(form),
                    exc.fields,
                    event=ev,
                    skips=repo.list_skips(conn, event_id),
                )
                return _html(request, "event_form.html", 422, back=back, **ctx)
            except EditConflict:
                # Keep what the person typed; the form now carries the newer rev, so pressing
                # Save again knowingly replaces the other change.
                ev = repo.get_event(conn, event_id)
                yours = forms.values_from_form(form)
                ctx = views.form_context(
                    conn,
                    yours,
                    event=ev,
                    skips=repo.list_skips(conn, event_id),
                    conflict=views.conflict_context(conn, ev, yours),
                )
                return _html(request, "event_form.html", 409, back=back, **ctx)
            return _see_other(_return_to(back, saved=f"Saved “{ev.title}”."))

    return await run_in_threadpool(work)


@router.post("/events/{event_id}/delete")
async def delete(event_id: int, request: Request):
    form = forms.parse_body(await request.body())
    back = _back(form)

    def work() -> Response:
        with _open(request) as conn:
            try:
                repo.soft_delete(conn, event_id, _rev(form))
            except EditConflict:
                ev = repo.get_event(conn, event_id)
                ctx = views.form_context(
                    conn,
                    forms.event_to_values(ev),
                    event=ev,
                    skips=repo.list_skips(conn, event_id),
                    error_message="Someone else changed this event, "
                    "so it was not deleted. Check it and try again.",
                )
                return _html(request, "event_form.html", 409, back=back, **ctx)
            return _see_other(_return_to(back, deleted=str(event_id)))

    return await run_in_threadpool(work)


@router.post("/events/{event_id}/restore")
def restore(event_id: int, conn=Depends(get_conn)):
    ev = repo.restore(conn, event_id)
    return _see_other("/?saved=" + _quote(f"Restored “{ev.title}”."))


@router.post("/events/{event_id}/skip")
async def skip(event_id: int, request: Request):
    """Skip one date, or every date from `date` through `last`. Returns to `back` if given."""
    form = forms.parse_body(await request.body())
    back = _back(form)

    def work() -> Response:
        with _open(request) as conn:
            first = (form.get("date") or [""])[0]
            last = (form.get("last") or [""])[0].strip()
            if last:
                n = repo.add_skip_range(conn, event_id, first, last)
                msg = f"Skipped {n} day{'' if n == 1 else 's'}."
            else:
                repo.add_skip(conn, event_id, first)
                msg = f"Day off on {first}."
        if back:
            return _see_other(_return_to(back, saved=msg))
        return _see_other(f"/events/{event_id}/edit")

    return await run_in_threadpool(work)


@router.post("/events/{event_id}/day")
async def day_note(event_id: int, request: Request):
    """Set or clear the in/out note on one day of a schedule, then return to where you were."""
    form = forms.parse_body(await request.body())
    back = _back(form)

    def work() -> Response:
        with _open(request) as conn:
            day = (form.get("date") or [""])[0]
            note = (form.get("note") or [""])[0]
            repo.set_day_note(conn, event_id, day, note or None)
            ev = repo.get_event(conn, event_id)
        msg = f"{ev.title} on {day}: {note}." if note else f"{ev.title} on {day}: not decided."
        if back:
            return _see_other(_return_to(back, saved=msg))
        return _see_other(f"/events/{event_id}/edit")

    return await run_in_threadpool(work)


@router.post("/events/{event_id}/unskip")
async def unskip(event_id: int, request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            repo.get_event(conn, event_id)
            repo.remove_skip(conn, event_id, (form.get("date") or [""])[0])
        return _see_other(f"/events/{event_id}/edit")

    return await run_in_threadpool(work)


@router.get("/deleted")
def deleted(request: Request, conn=Depends(get_conn)):
    return _html(request, "deleted.html", events=repo.list_deleted(conn))


@router.get("/render/current.png")
def render_current(request: Request, conn=Depends(get_conn)):
    result = render_service.render_current(conn, now_utc(request), request.app.state.render_cache)
    return Response(
        result.png,
        media_type="image/png",
        headers={"Cache-Control": "no-store", "X-Render-Hash": result.hash},
    )


@router.get("/render/{name}")
def render_by_hash(name: str, request: Request):
    png = request.app.state.render_cache.get(name[:-4]) if _HASH.match(name) else None
    if png is None:
        raise NotFound("That image is no longer available.")
    return Response(
        png,
        media_type="image/png",
        headers={"Cache-Control": "public, max-age=31536000, immutable"},
    )
