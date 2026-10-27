"""App factory and startup checks (fonts, single worker, migrations)."""

from __future__ import annotations

import asyncio
import contextlib
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Response

from trmnl_calendar import __version__
from trmnl_calendar.config import Settings, load_settings
from trmnl_calendar.render.fontcheck import verify_fonts
from trmnl_calendar.storage import backup, db, migrate

log = logging.getLogger("trmnl_calendar")


def create_app(settings: Settings | None = None) -> FastAPI:
    settings = settings or load_settings()

    async def backup_loop() -> None:
        while True:
            age = backup.latest_age_hours(settings.backup_dir)
            if age is None or age >= settings.backup_interval_hours:
                try:
                    path = await asyncio.to_thread(
                        backup.make_backup,
                        settings.db_path,
                        settings.backup_dir,
                        keep=settings.backup_keep,
                    )
                    log.info("backup written: %s", path.name)
                except Exception:  # keep the loop alive; healthz shows the age
                    log.exception("backup failed")
            await asyncio.sleep(600)

    @asynccontextmanager
    async def lifespan(app: FastAPI):
        if settings.web_concurrency != 1:
            raise RuntimeError("WEB_CONCURRENCY must be 1: one process, one worker (ADR-0007)")
        verify_fonts()
        conn = db.connect(settings.db_path)
        applied = migrate.migrate(conn, settings.db_path, settings.backup_dir)
        log.info("migrations applied: %s", applied or "none")
        app.state.conn = conn
        task = asyncio.create_task(backup_loop())
        try:
            yield
        finally:
            task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await task
            conn.close()

    app = FastAPI(title="TRMNL Calendar", version=__version__, lifespan=lifespan)
    app.state.settings = settings

    @app.get("/healthz")
    def healthz(response: Response) -> dict:
        body: dict = {"status": "ok", "version": __version__}
        try:
            body["schema_version"] = migrate.current_version(app.state.conn)
        except Exception as exc:
            response.status_code = 503
            return {"status": "error", "version": __version__, "detail": type(exc).__name__}
        body["last_backup_age_hours"] = backup.latest_age_hours(settings.backup_dir)
        body["scheduler"] = "not_started"  # slice 4
        return body

    return app


def app_factory() -> FastAPI:
    logging.basicConfig(level=logging.INFO)
    return create_app()
