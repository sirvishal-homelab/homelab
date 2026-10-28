"""What the status panel and /display/status show. No secrets, ever."""

from __future__ import annotations

import sqlite3
from datetime import datetime

from trmnl_calendar.config import Settings
from trmnl_calendar.storage.repositories import push_state as state_repo
from trmnl_calendar.storage.repositories import views as views_repo


def screen_status(conn: sqlite3.Connection, settings: Settings, scheduler) -> dict:
    st = state_repo.get_state(conn)
    display = views_repo.get_display(conn)
    nxt: datetime | None = getattr(scheduler, "next_check_at", None)
    return {
        "configured": settings.terminus_configured,
        "enabled": display.enabled,
        "status": st.last_status,
        "last_error": st.last_error,
        "last_attempt_at": st.last_attempt_at,
        "last_pushed_at": st.last_pushed_at,
        "image_hash": st.last_hash,
        "screen_id": st.screen_id,
        "failures": st.failures,
        "next_check_at": nxt.strftime("%Y-%m-%dT%H:%M:%SZ") if nxt else None,
        "interval_seconds": settings.push_interval_seconds,
        "device_refresh_seconds": state_repo.device_refresh_seconds(conn),
        "model_name": settings.terminus_model_name,
        "terminus_url": settings.terminus_url or None,
        "public_base_url": settings.public_base_url or None,
        "active_view": {
            "id": views_repo.active_view(conn).id,
            "name": views_repo.active_view(conn).name,
        },
    }
