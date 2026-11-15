"""The one background loop (docs/05 section 3).

Every `poll_s`: read Beszel, evaluate, render, and upload a frame only when it changed.
While an alert is active, `status.jpg` (overview) and `alert.jpg` (card) are both stored on the
screen and the loop switches between them every `alert_flip_s` with `/set?img=` only: no upload.
If Beszel fails twice in a row, the overview is replaced by an amber "Beszel unreachable" card,
red after `unreachable_grace_min`. If Beszel's data has an unknown shape, a "Beszel data changed"
card is shown instead of numbers that could be wrong (F-16).
"""

from __future__ import annotations

import asyncio
import contextlib
import logging
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import UTC, datetime
from zoneinfo import ZoneInfo

import httpx

from esp_screen import publish
from esp_screen.beszel.client import BeszelClient, BeszelDataError, BeszelError
from esp_screen.config import Settings
from esp_screen.device.client import DeviceError, ScreenClient
from esp_screen.domain.alerts import Card, alerts_for, card_for, track_since
from esp_screen.domain.evaluate import ContainerMemory, Thresholds, evaluate, remember
from esp_screen.domain.model import Status, SystemState
from esp_screen.render import frames
from esp_screen.storage import db, repo

log = logging.getLogger("esp_screen")

ALERT_IMAGE = "alert.jpg"
FAILURES_BEFORE_CARD = 2  # one failed read is not worth a warning


@dataclass
class BeszelHealth:
    last_ok_at: str | None = None
    last_error: str | None = None
    data_changed: bool = False
    failures: int = 0
    failing_since: datetime | None = None


@dataclass
class ScreenPlan:
    """What is stored on the screen and whether the loop flips between the two images."""

    alert_hash: str | None = None
    flipping: bool = False
    showing_alert: bool = False
    since: dict[str, datetime] = field(default_factory=dict)
    card: Card | None = None


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
        self.plan = ScreenPlan()
        self.last_status: Status | None = None
        self.last_states: dict[str, str] = {}  # system id -> state, for the settings page
        self.last_overview: bytes | None = None  # for the settings page preview
        self.last_alert: bytes | None = None
        self._last_good: datetime | None = None
        self._wake: asyncio.Event | None = None
        self._beszel = BeszelClient(
            settings.beszel_url,
            settings.beszel_user,
            settings.beszel_password,
            transport=beszel_transport,
            now=now,
        )

    async def aclose(self) -> None:
        await self._beszel.aclose()

    # --- helpers ----------------------------------------------------------------------------
    def _stamp(self, when: datetime) -> str:
        return when.astimezone(ZoneInfo(self.settings.timezone)).strftime("%H:%M")  # 24-hour

    def _minutes_since(self, when: datetime) -> int:
        return max(0, int((self.now() - when).total_seconds() // 60))

    def _setting(self, key: str, default):
        with contextlib.closing(db.connect(self.settings.db_path)) as conn:
            return repo.get_setting(conn, key, default)

    async def _screen(self, conn, action) -> bool:
        """Run `action(screen)` against the configured screen; record failures."""
        url = repo.get_setting(conn, "device_url", "")
        if not url:
            return False
        try:
            async with ScreenClient(url) as screen:
                await action(screen)
        except DeviceError as exc:
            log.warning("screen request failed: %s", exc)
            repo.record_push(conn, ok=False, frame_hash=None, error=str(exc))
            return False
        return True

    async def _upload_status(self, conn, data: bytes, *, show: bool) -> bool:
        """Upload status.jpg if it changed. Returns True if uploaded."""
        digest = frames.frame_hash(data)
        if digest == repo.get_push_state(conn).last_hash:
            return False

        async def action(screen):
            await screen.upload(publish.STATUS_IMAGE, data)
            if show:
                await screen.show(publish.STATUS_IMAGE)

        if not await self._screen(conn, action):
            return False
        repo.record_push(conn, ok=True, frame_hash=digest)
        return True

    async def _upload_alert(self, conn, data: bytes) -> None:
        digest = frames.frame_hash(data)
        if digest == self.plan.alert_hash:
            return
        if await self._screen(conn, lambda screen: screen.upload(ALERT_IMAGE, data)):
            self.plan.alert_hash = digest

    # --- one cycle ----------------------------------------------------------------------------
    async def tick(self) -> bool:
        """One cycle. Returns True if status.jpg was uploaded. Never raises."""
        try:
            snap = await self._beszel.snapshot()
        except BeszelError as exc:
            return await self._on_failure(exc)
        h = self.health
        h.last_ok_at = snap.taken_at.strftime("%Y-%m-%dT%H:%M:%SZ")
        h.last_error, h.data_changed, h.failures, h.failing_since = None, False, 0, None
        self._last_good = snap.taken_at
        self.last_states = {s.id: s.state.value for s in snap.systems}

        with contextlib.closing(db.connect(self.settings.db_path)) as conn:
            repo.sync_systems(conn, [(s.id, s.name) for s in snap.systems])
            if repo.seed_focus_by_name(conn, self.settings.focus_systems):
                log.info("focus systems moved from FOCUS_SYSTEMS (.env) to the settings page")
            focus = repo.focus_ids(conn)
            memory, allowed = repo.load_container_memory(conn)
            not_up = {s.id for s in snap.systems if s.state is not SystemState.UP}
            nxt = remember(ContainerMemory(memory, allowed), snap.containers, not_up)
            repo.save_container_memory(conn, nxt)
            status = evaluate(
                snap,
                focus,
                nxt,
                allowed_to_stop=allowed,
                labels=repo.labels(conn),
                ignore_alerts=repo.ignored_systems(conn),
                missing_limit=int(repo.get_setting(conn, "container_missing_polls", 2)),
                thresholds=Thresholds(
                    float(repo.get_setting(conn, "warn_pct", 75)),
                    float(repo.get_setting(conn, "crit_pct", 90)),
                ),
            )
            self.last_status = status
            alerts = alerts_for(status)
            self.plan.since = track_since(alerts, self.plan.since, snap.taken_at)
            card = card_for(alerts, self.plan.since)
            self.plan.card = card

            back_to_overview = False
            if card is not None:
                image = frames.alert(
                    card.title,
                    card.detail,
                    self._stamp(card.since),
                    self._minutes_since(card.since),
                    card.also,
                )
                self.last_alert = frames.encode_jpeg(image)
                await self._upload_alert(conn, self.last_alert)
                self.plan.flipping = self.plan.alert_hash is not None
            else:
                self.last_alert = None
                if self.plan.flipping:
                    log.info("all clear")
                back_to_overview = self.plan.showing_alert
                self.plan.flipping = self.plan.showing_alert = False

            overview = frames.encode_jpeg(frames.overview(status, self._stamp(snap.taken_at)))
            self.last_overview = overview
            # While the alert card is on screen, store the new overview without switching to
            # it: the flip timer decides what shows.
            uploaded = await self._upload_status(conn, overview, show=not self.plan.showing_alert)
            if back_to_overview and not uploaded:
                await self._screen(conn, lambda s: s.show(publish.STATUS_IMAGE))
            return uploaded

    async def _on_failure(self, exc: BeszelError) -> bool:
        h = self.health
        h.last_error = str(exc)
        h.failures += 1
        h.failing_since = h.failing_since or self.now()
        if isinstance(exc, BeszelDataError):
            h.data_changed = True
            log.error("Beszel data changed: %s", exc)
            data = frames.encode_jpeg(frames.data_changed(str(exc)))
        else:
            log.warning("Beszel read failed (%d in a row): %s", h.failures, exc)
            if h.failures < FAILURES_BEFORE_CARD:
                return False
            grace = float(self._setting("unreachable_grace_min", 5))
            minutes = self._minutes_since(self._last_good or h.failing_since)
            data = frames.encode_jpeg(
                frames.unreachable(
                    self._stamp(self._last_good) if self._last_good else None,
                    minutes,
                    red=minutes >= grace,
                )
            )
        # No fresh data: one card, no flipping. Uploading with show=True also leaves alert.jpg.
        self.plan.flipping = self.plan.showing_alert = False
        self.plan.card = None
        self.last_overview, self.last_alert = data, None
        with contextlib.closing(db.connect(self.settings.db_path)) as conn:
            return await self._upload_status(conn, data, show=True)

    # --- flipping ---------------------------------------------------------------------------
    async def flip(self) -> None:
        """Switch between overview and alert card; /set only, no upload."""
        if not self.plan.flipping:
            return
        target = publish.STATUS_IMAGE if self.plan.showing_alert else ALERT_IMAGE
        with contextlib.closing(db.connect(self.settings.db_path)) as conn:
            if await self._screen(conn, lambda s: s.show(target)):
                self.plan.showing_alert = target == ALERT_IMAGE

    # --- settings page hooks ------------------------------------------------------------------
    def request_refresh(self, *, force: bool = False) -> None:
        """Run a cycle now (after a settings change). `force` re-uploads even if unchanged."""
        if force:
            with contextlib.closing(db.connect(self.settings.db_path)) as conn:
                repo.clear_push_hash(conn)
            self.plan.alert_hash = None
        if self._wake is not None:
            self._wake.set()

    # --- forever ----------------------------------------------------------------------------
    async def run_forever(self) -> None:
        self._wake = asyncio.Event()
        next_poll = time.monotonic()
        while True:
            try:
                if time.monotonic() >= next_poll:
                    await self.tick()
                    next_poll = time.monotonic() + max(15.0, float(self._setting("poll_s", 60)))
                elif self.plan.flipping:
                    await self.flip()
            except Exception:  # keep the loop alive whatever happens; /healthz shows state
                log.exception("status loop failed")
            wait = next_poll - time.monotonic()
            if self.plan.flipping:
                wait = min(wait, max(3.0, float(self._setting("alert_flip_s", 10))))
            try:
                await asyncio.wait_for(self._wake.wait(), timeout=max(0.5, wait))
            except TimeoutError:
                continue
            self._wake.clear()
            next_poll = time.monotonic()  # woken by the settings page: read now
