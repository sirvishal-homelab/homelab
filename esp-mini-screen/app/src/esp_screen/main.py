"""App factory and start-up (fonts, single worker, migrations, backups, first push)."""

from __future__ import annotations

import asyncio
import contextlib
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Response

from esp_screen import __version__, publish
from esp_screen.config import Settings, load_settings
from esp_screen.loop import Runner
from esp_screen.render.fonts import verify_fonts
from esp_screen.storage import backup, db, migrate, repo

log = logging.getLogger("esp_screen")


def create_app(settings: Settings | None = None, *, beszel_transport=None) -> FastAPI:
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
            except Exception:  # keep the loop alive; /healthz shows the age
                log.exception("backup failed")

    @asynccontextmanager
    async def lifespan(app: FastAPI):
        if settings.web_concurrency != 1:
            raise RuntimeError("WEB_CONCURRENCY must be 1: one process, one worker")
        verify_fonts()
        with contextlib.closing(db.connect(settings.db_path)) as conn:
            applied = migrate.migrate(conn, settings.db_path, settings.backup_dir)
            log.info("migrations applied: %s", applied or "none")
            repo.seed_device_url(conn, settings.screen_url)
        backup_if_due()
        runner = (
            Runner(settings, beszel_transport=beszel_transport)
            if settings.beszel_configured
            else None
        )
        app.state.runner = runner

        async def screen_loop() -> None:
            await publish.push_startup_frame(settings)
            if runner is None:
                log.info("Beszel is not configured (.env): showing the start-up frame only")
                return
            if not settings.focus_systems:
                log.info("no FOCUS_SYSTEMS set: the overview shows the counters only")
            await runner.run_forever()

        tasks = [asyncio.create_task(backup_loop()), asyncio.create_task(screen_loop())]
        app.state.tasks = tasks
        try:
            yield
        finally:
            for t in tasks:
                t.cancel()
            for t in tasks:
                with contextlib.suppress(asyncio.CancelledError):
                    await t
            if runner is not None:
                await runner.aclose()

    app = FastAPI(title="ESP Mini Screen", version=__version__, lifespan=lifespan)
    app.state.settings = settings
    app.state.runner = None

    @app.get("/healthz")
    def healthz(response: Response) -> dict:
        body: dict = {"status": "ok", "version": __version__}
        try:
            with contextlib.closing(db.connect(settings.db_path)) as conn:
                body["schema_version"] = migrate.current_version(conn)
                push = repo.get_push_state(conn)
                body["screen"] = {
                    "configured": bool(repo.get_setting(conn, "device_url", "")),
                    "last_push_at": push.last_push_at,
                    "last_result": push.last_result,
                    "last_error": push.last_error,
                }
        except Exception as exc:
            log.exception("healthz: database check failed")
            response.status_code = 503
            return {"status": "error", "version": __version__, "detail": type(exc).__name__}
        body["beszel_configured"] = settings.beszel_configured
        runner = app.state.runner
        if runner is not None:
            h = runner.health
            body["beszel"] = {
                "last_ok_at": h.last_ok_at,
                "last_error": h.last_error,
                "data_changed": h.data_changed,
            }
            st = runner.last_status
            if st is not None:
                body["overview"] = {
                    "focus": [f"{r.label}:{r.state.value}" for r in st.rows],
                    "containers": f"{st.containers.up}/{st.containers.total}",
                    "other_systems": f"{st.others.up}/{st.others.total}",
                    "paused": st.others.paused,
                }
        body["last_backup_age_hours"] = backup.latest_age_hours(settings.backup_dir)
        return body

    return app


def app_factory() -> FastAPI:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    return create_app()
