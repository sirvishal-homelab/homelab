"""Render the active view and push it to Terminus only when it changed (docs/02 s3, s4, s7)."""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable
from contextlib import closing
from dataclasses import dataclass
from datetime import UTC, datetime

from trmnl_calendar.config import Settings
from trmnl_calendar.publish.terminus import TerminusClient, TerminusError
from trmnl_calendar.render.pipeline import RenderCache
from trmnl_calendar.storage import db
from trmnl_calendar.storage.repositories import push_state as state_repo
from trmnl_calendar.storage.repositories import views as views_repo
from trmnl_calendar.web import render_service

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class PushOutcome:
    status: str  # ok | unchanged | error | disabled | not_configured
    message: str = ""
    image_hash: str | None = None
    screen_id: int | None = None


def _iso(now: datetime) -> str:
    return now.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


class Publisher:
    def __init__(
        self,
        settings: Settings,
        cache: RenderCache,
        clock: Callable[[], datetime],
        client: TerminusClient | None = None,
    ):
        self.settings = settings
        self.cache = cache
        self.clock = clock
        self._client = client
        self._lock = threading.Lock()  # one push at a time (scheduler and "Push now")

    @property
    def client(self) -> TerminusClient:
        if self._client is None:
            s = self.settings
            self._client = TerminusClient(
                s.terminus_url,
                s.terminus_login,
                s.terminus_password,
                verify_tls=s.terminus_verify_tls,
            )
        return self._client

    def image_url(self, image_hash: str) -> str:
        return f"{self.settings.public_base_url}/render/{image_hash}.png"

    def run(self, force: bool = False) -> PushOutcome:
        if not self.settings.terminus_configured:
            return PushOutcome("not_configured", "Terminus is not set up in .env yet.")
        with self._lock, closing(db.connect(self.settings.db_path)) as conn:
            return self._run(conn, force)

    def _run(self, conn, force: bool) -> PushOutcome:
        now = self.clock()
        display = views_repo.get_display(conn)
        if not display.enabled:
            return PushOutcome("disabled", "Pushing to the screen is switched off.")
        state = state_repo.get_state(conn)
        try:
            result = render_service.render_current(conn, now, self.cache)
        except Exception as exc:  # bad data must never reach the screen
            log.exception("render failed")
            return self._fail(conn, state, now, f"Could not draw the calendar: {exc}", None)
        if (
            not force
            and state.last_hash == result.hash
            and state.last_status in ("ok", "unchanged")
        ):
            state_repo.save(
                conn, last_status="unchanged", last_attempt_at=_iso(now), last_error=None
            )
            return PushOutcome("unchanged", "Nothing changed.", result.hash, state.screen_id)

        state_repo.save(conn, last_status="in_progress", last_attempt_at=_iso(now))
        try:
            model = self.client.find_model(self.settings.terminus_model_name)
            profile = (display.width, display.height, display.bit_depth)
            theirs = (model.get("width"), model.get("height"), model.get("bit_depth"))
            if profile != theirs:
                raise TerminusError(
                    "config",
                    f"Terminus model '{model.get('name')}' is {theirs[0]} x {theirs[1]}, "
                    f"{theirs[2]}-bit, but this app draws {profile[0]} x {profile[1]}, "
                    f"{profile[2]}-bit. Nothing was sent.",
                )
            screen_id = self._send(state.screen_id, model["id"], self.image_url(result.hash))
        except TerminusError as exc:
            return self._fail(conn, state, now, exc.message, result.hash)
        state_repo.save(
            conn,
            screen_id=screen_id,
            last_hash=result.hash,
            last_status="ok",
            last_error=None,
            last_attempt_at=_iso(now),
            last_pushed_at=_iso(now),
            failures=0,
        )
        log.info("pushed %s to Terminus screen %s", result.hash, screen_id)
        return PushOutcome("ok", "Sent to the screen.", result.hash, screen_id)

    def _send(self, screen_id: int | None, model_id: int, uri: str) -> int:
        """PATCH the known screen; if it is gone (404) find it by name or create it."""
        c, name = self.client, self.settings.terminus_screen_name
        if screen_id is not None:
            try:
                return c.patch_screen(screen_id, uri)["id"]
            except TerminusError as exc:
                if exc.kind != "not_found":
                    raise
        found = None
        try:
            found = next((s for s in c.screens() if s.get("name") == name), None)
        except TerminusError as exc:
            if exc.kind not in ("not_found", "rejected"):
                raise
        if found is not None and found.get("id") != screen_id:
            return c.patch_screen(found["id"], uri)["id"]
        return c.create_screen(model_id, name, "Calendar", uri)["id"]

    def _fail(
        self, conn, state, now: datetime, message: str, image_hash: str | None
    ) -> PushOutcome:
        log.warning("push failed: %s", message)
        state_repo.save(
            conn,
            last_status="error",
            last_error=message[:500],
            last_attempt_at=_iso(now),
            failures=state.failures + 1,
        )
        return PushOutcome("error", message, image_hash, state.screen_id)
