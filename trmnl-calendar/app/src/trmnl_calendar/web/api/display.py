"""Push status, Push now, the connection check and the on/off switch (docs/06 section 3)."""

from __future__ import annotations

import sqlite3
from typing import Any

from fastapi import APIRouter, Body, Depends, Request

from trmnl_calendar.errors import BadGateway, NotConfigured, ValidationFailed
from trmnl_calendar.publish import check, status
from trmnl_calendar.storage.repositories import push_state as state_repo
from trmnl_calendar.storage.repositories import views as views_repo
from trmnl_calendar.web.api.views import display_to_api
from trmnl_calendar.web.deps import get_conn

router = APIRouter(prefix="/api/v1", tags=["display"])


@router.patch("/display")
def patch_display(
    body: dict[str, Any] = Body(...), conn: sqlite3.Connection = Depends(get_conn)
) -> dict:
    unknown = set(body) - {"enabled"}
    if unknown:
        raise ValidationFailed({k: "Unknown field" for k in sorted(unknown)})
    if not isinstance(body.get("enabled"), bool):
        raise ValidationFailed({"enabled": "Must be true or false"})
    conn.execute("UPDATE displays SET enabled = ? WHERE id = ?", (int(body["enabled"]), 1))
    conn.commit()
    return display_to_api(views_repo.get_display(conn))


@router.get("/display/status")
def display_status(request: Request, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    return status.screen_status(conn, request.app.state.settings, request.app.state.scheduler)


@router.post("/display/push")
def push_now(request: Request, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    """Render and push even if nothing changed. 502 with Terminus's reason if it fails."""
    out = request.app.state.publisher.run(force=True)
    if out.status == "not_configured":
        raise NotConfigured(out.message)
    if out.status == "disabled":
        raise NotConfigured(out.message)
    if out.status == "error":
        raise BadGateway(out.message)
    return {"status": out.status, "image_hash": out.image_hash, "screen_id": out.screen_id}


@router.post("/display/check")
def connection_check(request: Request, conn: sqlite3.Connection = Depends(get_conn)) -> dict:
    steps = check.run_check(
        request.app.state.settings,
        views_repo.get_display(conn),
        state_repo.get_state(conn).screen_id,
    )
    return {
        "ok": not any(s.status == "fail" for s in steps),
        "steps": [s.__dict__ for s in steps],
    }
