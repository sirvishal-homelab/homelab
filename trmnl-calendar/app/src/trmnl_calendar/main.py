"""App factory and startup checks (fonts, single worker, migrations)."""

from __future__ import annotations

import asyncio
import contextlib
import logging
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Response
from fastapi.staticfiles import StaticFiles

from trmnl_calendar import __version__
from trmnl_calendar.config import Settings, load_settings
from trmnl_calendar.holiday_source import HolidaySource
from trmnl_calendar.publish.publisher import Publisher
from trmnl_calendar.publish.scheduler import Scheduler
from trmnl_calendar.render.fontcheck import verify_fonts
from trmnl_calendar.render.pipeline import RenderCache
from trmnl_calendar.storage import backup, db, migrate
from trmnl_calendar.web import (
    errors,
    pages,
    pages_holidays,
    pages_manage,
    pages_views,
    render_service,
)
from trmnl_calendar.web.api import display as display_api
from trmnl_calendar.web.api import events as events_api
from trmnl_calendar.web.api import export as export_api
from trmnl_calendar.web.api import holidays as holidays_api
from trmnl_calendar.web.api import manage as manage_api
from trmnl_calendar.web.api import views as views_api

log = logging.getLogger("trmnl_calendar")


# Writes that cannot change the picture: previews, the who-am-I cookie, the push/check actions.
NO_PUSH_PATHS = {
    "/api/v1/views/preview",
    "/api/v1/me",
    "/who",
    "/api/v1/display/push",
    "/api/v1/display/check",
    "/settings/screen/push",
    "/settings/screen/check",
}


def create_app(settings: Settings | None = None, clock=render_service.now_utc) -> FastAPI:
    settings = settings or load_settings()

    def backup_if_due() -> None:
        age = backup.latest_age_hours(settings.backup_dir)
        if age is None or age >= settings.backup_interval_hours:
            path = backup.make_backup(
                settings.db_path, settings.backup_dir, keep=settings.backup_keep
            )
            log.info("backup written: %s", path.name)

    async def backup_loop() -> None:
        while True:
            await asyncio.sleep(600)
            try:
                await asyncio.to_thread(backup_if_due)
            except Exception:  # keep the loop alive; healthz shows the age
                log.exception("backup failed")

    @asynccontextmanager
    async def lifespan(app: FastAPI):
        if settings.web_concurrency != 1:
            raise RuntimeError("WEB_CONCURRENCY must be 1: one process, one worker (ADR-0007)")
        verify_fonts()
        conn = db.connect(settings.db_path)
        applied = migrate.migrate(conn, settings.db_path, settings.backup_dir)
        log.info("migrations applied: %s", applied or "none")
        conn.close()  # request handlers open their own short-lived connection (thread safety)
        backup_if_due()  # synchronous, so a backup exists before the first request
        task = asyncio.create_task(backup_loop())
        scheduler = app.state.scheduler
        if scheduler is not None:
            scheduler.start()
        else:
            log.info("Terminus is not configured (.env): pushing to the screen is off")
        try:
            yield
        finally:
            if scheduler is not None:
                await scheduler.stop()
            task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await task

    app = FastAPI(title="TRMNL Calendar", version=__version__, lifespan=lifespan)
    app.state.settings = settings
    app.state.clock = clock
    app.state.render_cache = RenderCache()
    app.state.holiday_source = HolidaySource(settings.holiday_api_url)
    app.state.publisher = Publisher(settings, app.state.render_cache, clock)
    app.state.scheduler = (
        Scheduler(
            app.state.publisher,
            settings.push_interval_seconds,
            settings.push_debounce_seconds,
            settings.push_startup_delay_seconds,
        )
        if settings.terminus_configured
        else None
    )

    @app.middleware("http")
    async def push_after_changes(request, call_next):
        response = await call_next(request)
        scheduler = request.app.state.scheduler
        if (
            scheduler is not None
            and request.method in ("POST", "PUT", "PATCH", "DELETE")
            and response.status_code < 400
            and request.url.path not in NO_PUSH_PATHS
        ):
            scheduler.request_push("change")
        return response

    errors.install(app)
    app.include_router(events_api.router)
    app.include_router(manage_api.router)
    app.include_router(views_api.router)
    app.include_router(display_api.router)
    app.include_router(export_api.router)
    app.include_router(holidays_api.router)
    app.include_router(pages_manage.router)
    app.include_router(pages_views.router)
    app.include_router(pages_holidays.router)
    app.include_router(pages.router)
    app.mount(
        "/static", StaticFiles(directory=Path(__file__).parent / "web" / "static"), name="static"
    )

    @app.get("/healthz")
    def healthz(response: Response) -> dict:
        body: dict = {"status": "ok", "version": __version__}
        try:
            with contextlib.closing(db.connect(settings.db_path)) as conn:
                body["schema_version"] = migrate.current_version(conn)
        except Exception as exc:
            log.exception("healthz: database check failed")
            response.status_code = 503
            return {"status": "error", "version": __version__, "detail": type(exc).__name__}
        body["last_backup_age_hours"] = backup.latest_age_hours(settings.backup_dir)
        sched = app.state.scheduler
        body["scheduler"] = "off" if sched is None else ("running" if sched.alive else "stopped")
        return body

    return app


def app_factory() -> FastAPI:
    logging.basicConfig(level=logging.INFO)
    return create_app()
