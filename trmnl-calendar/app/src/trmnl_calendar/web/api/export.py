"""`GET /api/v1/export`: all calendar data as a JSON download (docs/06)."""

from __future__ import annotations

import json
import sqlite3
from datetime import datetime

from fastapi import APIRouter, Depends, Request, Response

from trmnl_calendar import __version__
from trmnl_calendar.storage.repositories import export as export_repo
from trmnl_calendar.web.deps import get_conn, now_utc

router = APIRouter(prefix="/api/v1", tags=["export"])


@router.get("/export")
def export(request: Request, conn: sqlite3.Connection = Depends(get_conn)) -> Response:
    now: datetime = now_utc(request)
    body = export_repo.export_all(conn, __version__, now)
    name = f"calendar-export-{now.date().isoformat()}.json"
    return Response(
        json.dumps(body, indent=2, ensure_ascii=False),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="{name}"'},
    )
