"""One loop: push shortly after a change, and re-check on a timer (docs/02 section 4).

A change wakes the loop; it waits a few seconds so a burst of edits becomes one push, then renders.
The Publisher itself skips the network call when the picture has not changed.
"""

from __future__ import annotations

import asyncio
import contextlib
import logging
from datetime import UTC, datetime, timedelta

from trmnl_calendar.publish.publisher import Publisher

log = logging.getLogger(__name__)


class Scheduler:
    def __init__(
        self, publisher: Publisher, interval: int, debounce: float, startup_delay: float = 10.0
    ):
        self.startup_delay = startup_delay
        self.publisher = publisher
        self.interval = interval
        self.debounce = debounce
        self.next_check_at: datetime | None = None
        self.alive = False
        self._wake: asyncio.Event | None = None
        self._loop: asyncio.AbstractEventLoop | None = None
        self._task: asyncio.Task | None = None

    def start(self) -> None:
        self._loop = asyncio.get_running_loop()
        self._wake = asyncio.Event()
        self._task = asyncio.create_task(self._run())

    async def stop(self) -> None:
        if self._task:
            self._task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await self._task
        self.alive = False

    def request_push(self, reason: str = "change") -> None:
        """Safe to call from any thread."""
        if self._loop is None or self._wake is None:
            return
        with contextlib.suppress(RuntimeError):  # loop closed during shutdown
            self._loop.call_soon_threadsafe(self._wake.set)

    async def _run(self) -> None:
        assert self._wake is not None
        self.alive = True
        first = True
        while True:
            if first:
                # Terminus fetches the image from this app during the push, so wait until the
                # web server is accepting requests.
                await asyncio.sleep(self.startup_delay)
            else:
                self.next_check_at = datetime.now(UTC) + timedelta(seconds=self.interval)
                with contextlib.suppress(TimeoutError):
                    await asyncio.wait_for(self._wake.wait(), timeout=self.interval)
            if self._wake.is_set():
                await asyncio.sleep(self.debounce)
                self._wake.clear()
            first = False
            try:
                await asyncio.to_thread(self.publisher.run)
            except Exception:  # keep the loop alive; the status panel shows the last error
                log.exception("scheduled push failed")
