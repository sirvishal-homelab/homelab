"""View pages: list, templates, editor with a live native-size preview (docs/17 section 1)."""

from __future__ import annotations

from contextlib import closing
from urllib.parse import quote

from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from starlette.concurrency import run_in_threadpool

from trmnl_calendar.errors import InUse, ValidationFailed
from trmnl_calendar.storage import db
from trmnl_calendar.storage.repositories import views as repo
from trmnl_calendar.web import forms, views
from trmnl_calendar.web.deps import chrome, get_conn, now_utc

router = APIRouter(include_in_schema=False)


def _page(request: Request, name: str, status: int = 200, **ctx) -> HTMLResponse:
    return HTMLResponse(
        views.render_page(name, **{**chrome(request), "nav": "views", **ctx}), status_code=status
    )


def _go(url: str) -> RedirectResponse:
    return RedirectResponse(url, status_code=303)


def _open(request: Request):
    return closing(db.connect(request.app.state.settings.db_path))


def _editor(request: Request, conn, values, view_id=None, status=200, **kw) -> HTMLResponse:
    ctx = views.view_editor_context(
        conn, now_utc(request), request.app.state.render_cache, values, view_id, **kw
    )
    return _page(request, "view_edit.html", status, **ctx)


@router.get("/views")
def views_list(request: Request, saved: str | None = None, conn=Depends(get_conn)):
    ctx = views.views_list_context(conn)
    return _page(request, "views.html", flash=(saved or "")[:80], **ctx)


@router.get("/views/new")
def view_new(request: Request, template: str = "blank", conn=Depends(get_conn)):
    try:
        values = repo.template_values(template)
    except ValidationFailed:
        values = repo.template_values("blank")
    values["options"] = {**repo.registry.default_options(values["layout"]), **values["options"]}
    return _editor(request, conn, values)


@router.post("/views")
async def view_create(request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        data = forms.view_form_to_dict(form)
        with _open(request) as conn:
            if (form.get("action") or [""])[0] == "preview":
                return _editor(request, conn, data)
            try:
                v = repo.create_view(conn, data["template"], data)
            except ValidationFailed as exc:
                return _editor(request, conn, data, status=422, errors=exc.fields)
            if (form.get("action") or [""])[0] == "save_show":
                repo.activate_view(conn, v.id)
            return _go(f"/views/{v.id}/edit?saved=1")

    return await run_in_threadpool(work)


@router.get("/views/{view_id}/edit")
def view_edit(view_id: int, request: Request, saved: str | None = None, conn=Depends(get_conn)):
    v = repo.get_view(conn, view_id)
    return _editor(request, conn, repo.view_to_dict(v), view_id, flash="Saved." if saved else "")


@router.post("/views/{view_id}/edit")
async def view_save(view_id: int, request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        data = forms.view_form_to_dict(form)
        action = (form.get("action") or [""])[0]
        with _open(request) as conn:
            repo.get_view(conn, view_id)
            if action == "preview":
                return _editor(request, conn, data, view_id)
            try:
                repo.update_view(conn, view_id, data)
            except ValidationFailed as exc:
                return _editor(request, conn, data, view_id, status=422, errors=exc.fields)
            if action == "save_show":
                repo.activate_view(conn, view_id)
            return _go(f"/views/{view_id}/edit?saved=1")

    return await run_in_threadpool(work)


@router.post("/views/{view_id}/activate")
async def view_activate(view_id: int, request: Request):
    def work() -> Response:
        with _open(request) as conn:
            v = repo.activate_view(conn, view_id)
            return _go("/views?saved=" + quote(f"Now showing “{v.name}”."))

    return await run_in_threadpool(work)


@router.post("/views/{view_id}/delete")
async def view_delete(view_id: int, request: Request):
    def work() -> Response:
        with _open(request) as conn:
            try:
                repo.delete_view(conn, view_id)
            except InUse as exc:
                v = repo.get_view(conn, view_id)
                return _editor(
                    request,
                    conn,
                    repo.view_to_dict(v),
                    view_id,
                    status=409,
                    error_message=exc.message,
                )
            return _go("/views?saved=View%20deleted.")

    return await run_in_threadpool(work)


@router.post("/views/{view_id}/move")
async def view_move(view_id: int, request: Request):
    form = forms.parse_body(await request.body())

    def work() -> Response:
        with _open(request) as conn:
            ids = [v.id for v in repo.list_views(conn)]
            if view_id in ids:
                i = ids.index(view_id)
                j = i + (-1 if (form.get("direction") or ["1"])[0] == "-1" else 1)
                if 0 <= j < len(ids):
                    ids[i], ids[j] = ids[j], ids[i]
                    repo.reorder_views(conn, ids)
            return _go("/views")

    return await run_in_threadpool(work)
