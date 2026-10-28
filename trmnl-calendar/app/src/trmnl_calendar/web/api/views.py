"""Views, layouts, templates, preview and the display (docs/06 section 3)."""

from __future__ import annotations

import sqlite3
from datetime import date
from typing import Any

from fastapi import APIRouter, Body, Depends, Request
from fastapi.responses import JSONResponse

from trmnl_calendar.domain.models import Display, View
from trmnl_calendar.domain.view_templates import TEMPLATES
from trmnl_calendar.errors import ValidationFailed
from trmnl_calendar.render.layouts import registry
from trmnl_calendar.storage.repositories import views as repo
from trmnl_calendar.web import render_service
from trmnl_calendar.web.deps import get_conn, now_utc

router = APIRouter(prefix="/api/v1", tags=["views"])


def view_to_api(v: View, active_id: int | None) -> dict:
    return {
        "id": v.id,
        **repo.view_to_dict(v),
        "layout_label": registry.LAYOUTS[v.layout].LABEL,
        "position": v.position,
        "active": v.id == active_id,
    }


def display_to_api(d: Display) -> dict:
    return {
        "model_name": d.model_name,
        "width": d.width,
        "height": d.height,
        "bit_depth": d.bit_depth,
        "active_view_id": d.active_view_id,
        "enabled": d.enabled,
    }


def _active_id(conn: sqlite3.Connection) -> int | None:
    return repo.active_view(conn).id


@router.get("/layouts")
def list_layouts() -> list[dict]:
    return [
        {
            "id": lid,
            "label": m.LABEL,
            "options": [
                {
                    "key": k,
                    "label": registry.OPTION_LABELS[k],
                    "default": registry.OPTION_DEFAULTS[k],
                }
                for k in m.OPTIONS
            ],
        }
        for lid, m in registry.LAYOUTS.items()
    ]


@router.get("/templates")
def list_templates() -> list[dict]:
    return [
        {"id": tid, "label": t["label"], "help": t["help"], "layout": t["values"]["layout"]}
        for tid, t in TEMPLATES.items()
    ]


@router.get("/display")
def get_display(conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    return display_to_api(repo.get_display(conn))


@router.get("/views")
def list_views(conn: sqlite3.Connection = Depends(get_conn)) -> list[dict]:
    active = _active_id(conn)
    return [view_to_api(v, active) for v in repo.list_views(conn)]


@router.post("/views", status_code=201)
def create_view(
    body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> JSONResponse:
    unknown = set(body) - {"template", "name", "overrides"}
    if unknown:
        raise ValidationFailed({k: "Unknown field" for k in sorted(unknown)})
    overrides = body.get("overrides") or {}
    if not isinstance(overrides, dict):
        raise ValidationFailed({"overrides": "Send overrides as an object"})
    if "name" in body:
        overrides = {**overrides, "name": body["name"]}
    v = repo.create_view(conn, body.get("template", "blank"), overrides)
    return JSONResponse(view_to_api(v, _active_id(conn)), status_code=201)


@router.put("/views/order")
def order_views(
    body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> list[dict]:
    ids = body.get("ids")
    if not isinstance(ids, list) or not all(isinstance(i, int) for i in ids):
        raise ValidationFailed({"ids": "Send a list of view ids"})
    repo.reorder_views(conn, ids)
    active = _active_id(conn)
    return [view_to_api(v, active) for v in repo.list_views(conn)]


@router.post("/views/preview")
def preview(
    request: Request,
    body: dict[str, Any] = Body(...),
    conn: sqlite3.Connection = Depends(get_conn),
) -> dict:
    """Render unsaved settings. Stores nothing, pushes nothing (docs/17 section 1)."""
    data = body.get("view")
    if not isinstance(data, dict):
        raise ValidationFailed({"view": "Send the view settings"})
    as_of = None
    if body.get("as_of"):
        try:
            as_of = date.fromisoformat(str(body["as_of"]))
        except ValueError:
            raise ValidationFailed({"as_of": "Use a date like 2026-10-25"}) from None
    view = repo.build_view(conn, {**data, "name": "Preview"}, view_id=None)
    result = render_service.render_view(
        conn, view, now_utc(request), request.app.state.render_cache, as_of
    )
    return {
        "image": {
            "url": f"/render/{result.hash}.png",
            "width": result.width,
            "height": result.height,
            "bit_depth": result.bit_depth,
        },
        "ok": result.fit.ok,
        "fit": [
            {"code": f.code, "message": f.message, "where": f.where} for f in result.fit.findings
        ],
    }


@router.get("/views/{view_id}")
def get_view(view_id: int, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    return view_to_api(repo.get_view(conn, view_id), _active_id(conn))


@router.patch("/views/{view_id}")
def patch_view(
    view_id: int, body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> dict:
    return view_to_api(repo.update_view(conn, view_id, body), _active_id(conn))


@router.delete("/views/{view_id}")
def delete_view(view_id: int, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    repo.delete_view(conn, view_id)
    return {"deleted": view_id}


@router.post("/views/{view_id}/activate")
def activate(view_id: int, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    v = repo.activate_view(conn, view_id)
    return view_to_api(v, v.id)
