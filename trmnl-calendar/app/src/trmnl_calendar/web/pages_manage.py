"""Manage pages: categories, people, settings, who-am-I. Server-rendered, plain forms."""

from __future__ import annotations

from contextlib import closing

from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from starlette.concurrency import run_in_threadpool

from trmnl_calendar.errors import AppError, InUse, ValidationFailed
from trmnl_calendar.storage import db
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import people as ppl
from trmnl_calendar.storage.repositories import sections as sec
from trmnl_calendar.web import forms, views
from trmnl_calendar.web.deps import WHO_COOKIE, chrome, get_conn, safe_next, who_from_cookie

router = APIRouter(include_in_schema=False)


def _page(request: Request, name: str, status: int = 200, **ctx) -> HTMLResponse:
    return HTMLResponse(views.render_page(name, **{**chrome(request), **ctx}), status_code=status)


def _go(url: str) -> RedirectResponse:
    return RedirectResponse(url, status_code=303)


def _open(request: Request):
    return closing(db.connect(request.app.state.settings.db_path))


def _one(form: dict, key: str) -> str:
    return (form.get(key) or [""])[0].strip()


# ---- categories


@router.get("/categories")
def categories(request: Request, conn=Depends(get_conn)):
    return _page(
        request, "categories.html", nav="categories", **views.manage_categories_context(conn)
    )


@router.post("/categories")
async def category_add(request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            try:
                sec.create_section(
                    conn, _one(form, "name"), placement=_one(form, "placement") or "lane"
                )
            except ValidationFailed as exc:
                values = {
                    "name": _one(form, "name"),
                    "placement": _one(form, "placement") or "lane",
                }
                ctx = views.manage_categories_context(conn, values, exc.fields)
                return _page(request, "categories.html", 422, nav="categories", **ctx)
            return _go("/categories")

    return await run_in_threadpool(work)


@router.get("/categories/{section_id}/edit")
def category_edit_form(section_id: int, request: Request, conn=Depends(get_conn)):
    ctx = views.category_edit_context(conn, section_id)
    return _page(request, "category_edit.html", nav="categories", **ctx)


@router.post("/categories/{section_id}/edit")
async def category_edit(section_id: int, request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            patch = {
                "name": _one(form, "name"),
                "placement": _one(form, "placement"),
                "weight": _one(form, "weight"),
                "visible": "visible" in form,
            }
            try:
                sec.update_section(conn, section_id, patch)
            except ValidationFailed as exc:
                values = {**patch, "weight": patch["weight"]}
                ctx = views.category_edit_context(conn, section_id, values, exc.fields)
                return _page(request, "category_edit.html", 422, nav="categories", **ctx)
            return _go("/categories")

    return await run_in_threadpool(work)


@router.post("/categories/{section_id}/move")
async def category_move(section_id: int, request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            direction = -1 if _one(form, "direction") == "-1" else 1
            sec.move_section(conn, section_id, direction)
        return _go("/categories")

    return await run_in_threadpool(work)


@router.post("/categories/{section_id}/delete")
async def category_delete(section_id: int, request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            move_to = _one(form, "move_to")
            try:
                sec.delete_section(conn, section_id, int(move_to) if move_to.isdigit() else None)
            except InUse as exc:
                ctx = views.category_edit_context(conn, section_id, error_message=exc.message)
                return _page(request, "category_edit.html", 409, nav="categories", **ctx)
            return _go("/categories")

    return await run_in_threadpool(work)


# ---- people


@router.get("/people")
def people(request: Request, conn=Depends(get_conn)):
    return _page(request, "people.html", nav="people", **views.manage_people_context(conn))


@router.post("/people")
async def person_add(request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            try:
                ppl.create_person(conn, _one(form, "name"), _one(form, "label"))
            except ValidationFailed as exc:
                values = {"name": _one(form, "name"), "label": _one(form, "label")}
                ctx = views.manage_people_context(conn, values, exc.fields)
                return _page(request, "people.html", 422, nav="people", **ctx)
            return _go("/people")

    return await run_in_threadpool(work)


@router.get("/people/{person_id}/edit")
def person_edit_form(person_id: int, request: Request, conn=Depends(get_conn)):
    ctx = views.person_edit_context(conn, person_id)
    return _page(request, "person_edit.html", nav="people", **ctx)


@router.post("/people/{person_id}/edit")
async def person_edit(person_id: int, request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            patch = {
                "name": _one(form, "name"),
                "label": _one(form, "label"),
                "active": "active" in form,
            }
            try:
                ppl.update_person(conn, person_id, patch)
            except ValidationFailed as exc:
                ctx = views.person_edit_context(conn, person_id, patch, exc.fields)
                return _page(request, "person_edit.html", 422, nav="people", **ctx)
            return _go("/people")

    return await run_in_threadpool(work)


@router.post("/people/{person_id}/delete")
async def person_delete(person_id: int, request: Request):
    def work() -> Response:
        with _open(request) as conn:
            try:
                ppl.delete_person(conn, person_id)
            except InUse as exc:
                ctx = views.person_edit_context(conn, person_id, error_message=exc.message)
                return _page(request, "person_edit.html", 409, nav="people", **ctx)
            return _go("/people")

    return await run_in_threadpool(work)


# ---- settings


def _screen(request: Request, conn):
    return views.screen_context(conn, request.app.state.settings, request.app.state.scheduler)


@router.get("/settings")
def settings(request: Request, saved: int = 0, screen: str = "", conn=Depends(get_conn)):
    flash = {
        "pushed": "Sent to the screen.",
        "failed": "The push did not work. The reason is shown below.",
        "on": "Pushing is switched on.",
        "off": "Pushing is switched off.",
    }.get(screen, "")
    ctx = views.settings_context(
        conn, saved=bool(saved), screen=_screen(request, conn), screen_flash=flash
    )
    return _page(request, "settings.html", nav="settings", **ctx)


@router.post("/settings/screen/check")
async def screen_check(request: Request):
    def work() -> Response:
        from trmnl_calendar.publish import check
        from trmnl_calendar.storage.repositories import push_state as state_repo
        from trmnl_calendar.storage.repositories import views as views_repo

        with _open(request) as conn:
            steps = check.run_check(
                request.app.state.settings,
                views_repo.get_display(conn),
                state_repo.get_state(conn).screen_id,
                remember=lambda d, m: state_repo.remember_device(conn, d, m),
            )
            ctx = views.settings_context(conn, screen=_screen(request, conn), check_steps=steps)
            return _page(request, "settings.html", nav="settings", **ctx)

    return await run_in_threadpool(work)


@router.post("/settings/screen/push")
async def screen_push(request: Request):
    def work() -> Response:
        out = request.app.state.publisher.run(force=True)
        return _go("/settings?screen=" + ("pushed" if out.status == "ok" else "failed") + "#screen")

    return await run_in_threadpool(work)


@router.post("/settings/screen/enabled")
async def screen_enabled(request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        on = _one(form, "enabled") == "1"
        with _open(request) as conn:
            conn.execute("UPDATE displays SET enabled = ? WHERE id = 1", (int(on),))
            conn.commit()
        return _go("/settings?screen=" + ("on" if on else "off") + "#screen")

    return await run_in_threadpool(work)


@router.post("/settings")
async def settings_save(request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            patch = {k: _one(form, k) for k in ("name", "timezone", "week_start", "time_format")}
            try:
                cal.update_calendar(conn, patch)
            except ValidationFailed as exc:
                ctx = views.settings_context(conn, patch, exc.fields)
                return _page(request, "settings.html", 422, nav="settings", **ctx)
            return _go("/settings?saved=1")

    return await run_in_threadpool(work)


# ---- who am I (a convenience cookie, not authentication)


@router.get("/who")
def who(request: Request, next: str = "/", conn=Depends(get_conn)):
    me = who_from_cookie(request, conn)
    ctx = views.who_context(conn, me.id if me else None, safe_next(next))
    return _page(request, "who.html", **ctx)


@router.post("/who")
async def who_set(request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            raw = _one(form, "person_id")
            person_id = 0
            if raw.isdigit() and int(raw) > 0:
                try:
                    person_id = ppl.get_person(conn, int(raw)).id
                except AppError:
                    person_id = 0
            resp = _go(safe_next(_one(form, "next")))
            resp.set_cookie(
                WHO_COOKIE, str(person_id), max_age=365 * 24 * 3600, samesite="lax", httponly=True
            )
            return resp

    return await run_in_threadpool(work)
