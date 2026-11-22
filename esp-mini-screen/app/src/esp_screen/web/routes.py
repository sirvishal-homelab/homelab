"""Settings page routes (S-4). LAN only, no login (ADR-0003, N-8): it holds no secrets."""

from __future__ import annotations

import contextlib
import re

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response

from esp_screen import __version__, publish
from esp_screen.config import Settings
from esp_screen.device.client import DeviceError, ScreenClient
from esp_screen.render import frames
from esp_screen.storage import db, repo
from esp_screen.web import forms, page

_KEY = re.compile(r"^([A-Za-z0-9_-]+):([A-Za-z0-9][A-Za-z0-9_.-]*)$")


async def _screen_warnings(url: str) -> list[str]:
    if not url:
        return ["The screen address is not set: nothing is pushed until you set it below."]
    try:
        async with ScreenClient(url, timeout=2) as screen:
            album = await screen.album()
    except DeviceError:
        return []  # busy or briefly offline: the status box shows real push failures
    if album.autoplay:
        return [
            "The screen's album auto display is ON, so it rotates away from the status image. "
            "Turn off 'Enable image auto display' on the screen's Pictures page."
        ]
    return []


def build_router(settings: Settings, app_state) -> APIRouter:
    router = APIRouter()

    def runner():
        return app_state.runner

    def conn():
        return contextlib.closing(db.connect(settings.db_path))

    @router.get("/", response_class=HTMLResponse)
    async def settings_page(request: Request) -> HTMLResponse:
        q = request.query_params
        with conn() as c:
            url = repo.get_setting(c, "device_url", "")
        warnings = await _screen_warnings(url)
        with conn() as c:
            html = page.render(
                c,
                runner(),
                flags={k: k in q for k in ("saved", "pushed", "forgot")},
                warnings=warnings,
                tz=settings.timezone,
            )
        return HTMLResponse(html, headers={"Cache-Control": "no-store"})

    @router.post("/settings")
    async def save(request: Request) -> Response:
        form = await request.form()
        multi = {k: [str(v) for v in form.getlist(k)] for k in form.keys()}
        with conn() as c:
            known = {p.beszel_id for p in repo.list_systems(c)}
            parsed = forms.parse(multi, known)
            if not parsed.ok:
                html = page.render(c, runner(), errors=parsed.errors, tz=settings.timezone)
                return HTMLResponse(html, status_code=400)
            for key, value in parsed.numbers.items():
                repo.set_setting(c, key, value)
            repo.set_setting(c, "device_url", parsed.device_url)
            repo.set_setting(c, "show_updates", parsed.show_updates)
            repo.save_system_prefs(c, parsed.systems)
            repo.save_allowed_to_stop(c, parsed.allowed)
            repo.set_setting(c, "focus_seeded", True)  # the page now owns the focus systems
        r = runner()
        if r is not None:
            r.request_refresh()
        return RedirectResponse("/?saved=1", status_code=303)

    @router.post("/push")
    async def push_now() -> Response:
        r = runner()
        if r is not None:
            r.request_refresh(force=True)
        else:
            await publish.push_startup_frame(settings)
        return RedirectResponse("/?pushed=1", status_code=303)

    @router.post("/containers/forget")
    async def forget(request: Request) -> Response:
        form = await request.form()
        m = _KEY.match(str(form.get("key", "")))
        if m:
            with conn() as c:
                repo.forget_container(c, m.group(1), m.group(2))
            r = runner()
            if r is not None:
                r.request_refresh()
        return RedirectResponse("/?forgot=1", status_code=303)

    @router.get("/preview.jpg")
    async def preview(which: str = "overview") -> Response:
        r = runner()
        data = None
        if r is not None:
            data = r.last_alert if which == "alert" else r.last_overview
        if data is None:
            data = frames.encode_jpeg(frames.hello(__version__, "--:--"))
        return Response(data, media_type="image/jpeg", headers={"Cache-Control": "no-store"})

    return router
