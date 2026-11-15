"""The one background loop: read Beszel, evaluate, render, upload when the frame changed."""

from __future__ import annotations

import asyncio
import contextlib
import logging
from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, datetime
from zoneinfo import ZoneInfo

import httpx

from esp_screen import publish
from esp_screen.beszel.client import BeszelClient, BeszelDataError, BeszelError
from esp_screen.config import Settings
from esp_screen.device.client import DeviceError, ScreenClient
from esp_screen.domain.evaluate import ContainerMemory, Thresholds, evaluate, remember
from esp_screen.domain.model import Status, SystemState
from esp_screen.render import frames
from esp_screen.storage import db, repo

log = logging.getLogger("esp_screen")


@dataclass
class BeszelHealth:
    last_ok_at: str | None = None
    last_error: str | None = None
    data_changed: bool = False


class Runner:
    def __init__(
        self,
        settings: Settings,
        *,
        now: Callable[[], datetime] = lambda: datetime.now(UTC),
        beszel_transport: httpx.AsyncBaseTransport | None = None,
    ) -> None:
        self.settings = settings
        self.now = now
        self.health = BeszelHealth()
        self.last_status: Status | None = None
        self._beszel = BeszelClient(
            settings.beszel_url,
            settings.beszel_user,
            settings.beszel_password,
            transport=beszel_transport,
            now=now,
        )

    async def aclose(self) -> None:
        await self._beszel.aclose()

    def _stamp(self, when: datetime) -> str:
        return when.astimezone(ZoneInfo(self.settings.timezone)).strftime("%H:%M")  # 24-hour

    async def tick(self) -> bool:
        """One cycle. Returns True if a new frame was uploaded. Never raises."""
        try:
            snap = await self._beszel.snapshot()
        except BeszelDataError as exc:
            self.health.last_error, self.health.data_changed = str(exc), True
            log.error("Beszel data changed: %s", exc)  # S-3 shows this on the screen (F-16)
            return False
        except BeszelError as exc:
            self.health.last_error = str(exc)
            log.warning("Beszel read failed: %s", exc)  # S-3 adds the unreachable frame
            return False
        self.health.last_ok_at = snap.taken_at.strftime("%Y-%m-%dT%H:%M:%SZ")
        self.health.last_error, self.health.data_changed = None, False

        with contextlib.closing(db.connect(self.settings.db_path)) as conn:
            memory, allowed = repo.load_container_memory(conn)
            not_up = {s.id for s in snap.systems if s.state is not SystemState.UP}
            nxt = remember(ContainerMemory(memory, allowed), snap.containers, not_up)
            repo.save_container_memory(conn, nxt)
            status = evaluate(
                snap,
                self.settings.focus_systems,
                nxt,
                allowed_to_stop=allowed,
                missing_limit=int(repo.get_setting(conn, "container_missing_polls", 2)),
                thresholds=Thresholds(
                    float(repo.get_setting(conn, "warn_pct", 75)),
                    float(repo.get_setting(conn, "crit_pct", 90)),
                ),
            )
            self.last_status = status
            data = frames.encode_jpeg(frames.overview(status, self._stamp(snap.taken_at)))
            digest = frames.frame_hash(data)
            if digest == repo.get_push_state(conn).last_hash:
                return False
            url = repo.get_setting(conn, "device_url", "")
            if not url:
                return False
            try:
                async with ScreenClient(url) as screen:
                    await screen.push(publish.STATUS_IMAGE, data)
            except DeviceError as exc:
                log.warning("push to screen failed: %s", exc)
                repo.record_push(conn, ok=False, frame_hash=None, error=str(exc))
                return False
            repo.record_push(conn, ok=True, frame_hash=digest)
            return True

    def poll_seconds(self) -> float:
        with contextlib.closing(db.connect(self.settings.db_path)) as conn:
            return max(15.0, float(repo.get_setting(conn, "poll_s", 60)))

    async def run_forever(self) -> None:
        while True:
            try:
                await self.tick()
            except Exception:  # keep the loop alive whatever happens; /healthz shows state
                log.exception("status loop failed")
            await asyncio.sleep(self.poll_seconds())
