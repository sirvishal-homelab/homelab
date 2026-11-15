"""Glue between rendering, the screen client and the database (the only module with all three)."""

from __future__ import annotations

import asyncio
import contextlib
import logging
from collections.abc import Callable
from datetime import datetime
from zoneinfo import ZoneInfo

import httpx

from esp_screen import __version__
from esp_screen.config import Settings
from esp_screen.device.client import DeviceError, ScreenClient
from esp_screen.render import frames
from esp_screen.storage import db, repo

log = logging.getLogger("esp_screen")

STATUS_IMAGE = "status.jpg"  # one fixed name: the file is replaced, not added (R-4)


async def push_startup_frame(
    settings: Settings,
    *,
    now: Callable[[], datetime] = datetime.now,
    transport: httpx.AsyncBaseTransport | None = None,
) -> bool:
    """Render the start-up frame and show it. Never raises: the result goes to push_state."""
    with contextlib.closing(db.connect(settings.db_path)) as conn:
        url = repo.get_setting(conn, "device_url", "")
        if not url:
            log.info("screen address not set (SCREEN_URL or settings page): nothing pushed")
            return False
        stamp = now().astimezone(ZoneInfo(settings.timezone)).strftime("%H:%M")
        data = frames.encode_jpeg(frames.hello(__version__, stamp))
        digest = frames.frame_hash(data)
        async with ScreenClient(url, transport=transport) as screen:
            try:
                await screen.push(STATUS_IMAGE, data)
            except DeviceError as exc:
                log.warning("push to screen failed: %s", exc)
                repo.record_push(conn, ok=False, frame_hash=None, error=str(exc))
                return False
            repo.record_push(conn, ok=True, frame_hash=digest)
            log.info("start-up frame shown on the screen (%d bytes)", len(data))
            await warn_if_autoplay(screen)
        return True


async def warn_if_autoplay(screen: ScreenClient) -> None:
    """Advisory only: never turns a good push into a failure.

    Right after an upload the firmware can answer /album.json with a bare "OK" (seen on
    Ultra-V9.0.50, docs/03), so read it after a short pause and give up quietly if it is busy.
    """
    for delay in (1.0, 3.0):
        await asyncio.sleep(delay)
        try:
            album = await screen.album()
        except DeviceError as exc:
            log.debug("album settings not readable yet: %s", exc)
            continue
        if album.autoplay:
            log.warning(
                "the screen's album autoplay is ON: it will rotate away from the status image. "
                "Turn off 'Enable image auto display' on the screen's Pictures page."
            )
        return
    log.info("could not read the screen's album settings (it answered 'OK'); check skipped")
