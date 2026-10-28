import asyncio
import threading

from test_publisher import NOW, add, env  # noqa: F401  (the shared fixture)

from trmnl_calendar.publish.scheduler import Scheduler


def test_startup_push_then_debounced_change_push(env):  # noqa: F811
    fake, _, pub, conn, _ = env

    async def scenario():
        s = Scheduler(pub, interval=60, debounce=0.2, startup_delay=0)
        s.start()
        await asyncio.sleep(1.0)
        assert fake.calls("POST", "/api/screens") == 1  # the startup push
        # a burst of edits from request threads becomes one push
        for title in ("A", "B", "C"):
            add(conn, title)
            threading.Thread(target=s.request_push).start()
            await asyncio.sleep(0.02)
        await asyncio.sleep(1.2)
        assert fake.calls("PATCH", "/api/screens/") == 1
        assert s.alive and s.next_check_at is not None
        await s.stop()
        assert not s.alive

    asyncio.run(scenario())


def test_timer_check_pushes_nothing_when_unchanged(env):  # noqa: F811
    fake, _, pub, _, _ = env

    async def scenario():
        s = Scheduler(pub, interval=1, debounce=0.05, startup_delay=0)
        s.start()
        await asyncio.sleep(2.6)  # startup push, then at least two timer checks
        await s.stop()

    asyncio.run(scenario())
    assert fake.calls("POST", "/api/screens") == 1 and fake.calls("PATCH", "/api/screens/") == 0
