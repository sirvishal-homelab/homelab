import asyncio
import contextlib
from datetime import UTC, datetime, timedelta

from esp_screen.config import load_settings
from esp_screen.loop import Runner
from esp_screen.storage import db, migrate, repo
from tests.fakes.beszel import PASSWORD, USER, FakeHub
from tests.fakes.screen import FakeScreen


class Clock:
    def __init__(self):
        self.t = datetime(2026, 11, 15, 12, 30, tzinfo=UTC)  # 08:30 in Halifax

    def __call__(self):
        return self.t

    def advance(self, minutes=1):
        self.t += timedelta(minutes=minutes)


def setup(tmp_path, url, focus="node-a,node-b,nas"):
    s = load_settings(
        {
            "DATA_DIR": str(tmp_path),
            "SCREEN_URL": url,
            "BESZEL_URL": "http://beszel:8090",
            "BESZEL_USER": USER,
            "BESZEL_PASSWORD": PASSWORD,
            "FOCUS_SYSTEMS": focus,
        }
    )
    with contextlib.closing(db.connect(s.db_path)) as conn:
        migrate.migrate(conn, s.db_path, s.backup_dir)
        repo.seed_device_url(conn, s.screen_url)
    return s


def run_ticks(runner, clock, n, between=None):
    async def go():
        results = []
        for i in range(n):
            results.append(await runner.tick())
            if between:
                between(i)
        await runner.aclose()
        return results

    return asyncio.run(go())


def test_first_read_pushes_overview(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        assert run_ticks(runner, clock, 1) == [True]
    assert screen.shown == "/image/status.jpg"
    st = runner.last_status
    assert [r.label for r in st.rows] == ["node-a", "node-b", "nas"]
    assert (st.containers.up, st.containers.total) == (9, 9)
    assert (st.others.up, st.others.total, st.others.paused) == (3, 3, 1)
    assert runner.health.last_ok_at == "2026-11-15T12:30:00Z"


def test_same_minute_same_data_does_not_reupload(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        assert run_ticks(runner, clock, 2) == [True, False]
    assert sum(c.startswith("POST /doUpload") for c in screen.calls) == 1


def test_new_minute_uploads_again(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        assert run_ticks(runner, clock, 2, between=lambda i: clock.advance()) == [True, True]


def test_stopped_container_counts_after_two_reads(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()

    def world(i):
        clock.advance()
        if i == 0:
            hub.stop_container("docker-host", "app-a")

    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        run_ticks(runner, clock, 2, between=world)
        assert (runner.last_status.containers.up, runner.last_status.containers.total) == (9, 9)
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        run_ticks(runner, clock, 1)
    st = runner.last_status
    assert (st.containers.up, st.containers.total) == (8, 9)
    assert st.down_containers == ("docker-host/app-a",)


def test_beszel_down_keeps_last_frame_and_reports(tmp_path):
    hub, screen, clock = FakeHub(down=True), FakeScreen(), Clock()
    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        assert run_ticks(runner, clock, 1) == [False]
    assert screen.calls == []
    assert "cannot reach" in runner.health.last_error


def test_changed_beszel_data_is_flagged(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    del hub.systems["items"][0]["info"]
    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        run_ticks(runner, clock, 1)
    assert runner.health.data_changed is True


def test_screen_offline_is_recorded(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(drop_next=10), Clock()
    with screen.serve() as url:
        s = setup(tmp_path, url)
        runner = Runner(s, now=clock, beszel_transport=hub.transport)
        assert run_ticks(runner, clock, 1) == [False]
    with contextlib.closing(db.connect(s.db_path)) as conn:
        assert repo.get_push_state(conn).last_result == "error"


# --- S-3: alerts, flipping, Beszel failures ----------------------------------------------


def screen_calls(screen, kind):
    return [c for c in screen.calls if c.startswith(kind)]


def test_system_down_uploads_alert_and_flips(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    hub.set_status("node-b", "down")

    async def go(runner):
        await runner.tick()
        assert runner.plan.flipping
        await runner.flip()
        assert screen.shown == "/image/alert.jpg"
        await runner.flip()
        assert screen.shown == "/image/status.jpg"
        await runner.aclose()

    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        asyncio.run(go(runner))
    assert "/image/alert.jpg" in screen.files
    assert runner.plan.card.title == "node-b down"
    assert runner.last_status.down_systems == ("node-b",)


def test_flips_do_not_upload(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    hub.set_status("vm-1", "down")

    async def go(runner):
        await runner.tick()
        uploads = len(screen_calls(screen, "POST"))
        for _ in range(6):
            await runner.flip()
        assert len(screen_calls(screen, "POST")) == uploads
        await runner.aclose()

    with screen.serve() as url:
        asyncio.run(go(Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)))


def test_overview_refresh_while_alert_showing_does_not_switch(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    hub.set_status("node-b", "down")

    async def go(runner):
        await runner.tick()
        await runner.flip()  # alert card on screen
        clock.advance()
        assert await runner.tick() is True  # new minute: overview re-uploaded ...
        assert screen.shown == "/image/alert.jpg"  # ... but the card stays until the next flip
        await runner.aclose()

    with screen.serve() as url:
        asyncio.run(go(Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)))


def test_recovery_returns_to_overview(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    hub.set_status("node-b", "down")

    async def go(runner):
        await runner.tick()
        await runner.flip()
        hub.set_status("node-b", "up")
        await runner.tick()  # node-b is back: the overview changed, so it is re-uploaded
        assert not runner.plan.flipping
        assert screen.shown == "/image/status.jpg"
        assert runner.plan.card is None
        await runner.aclose()

    with screen.serve() as url:
        asyncio.run(go(Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)))


def test_paused_system_never_alerts(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    hub.set_status("node-b", "paused")
    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        run_ticks(runner, clock, 1)
    assert runner.plan.card is None
    assert "/image/alert.jpg" not in screen.files


def test_unreachable_card_after_two_failures_then_red(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()

    async def go(runner):
        await runner.tick()  # good read at 08:30
        hub.down = True
        clock.advance()
        assert await runner.tick() is False  # one failure: keep the overview
        clock.advance()
        assert await runner.tick() is True  # second: amber card
        amber = screen.files["/image/status.jpg"]
        clock.advance(5)
        assert await runner.tick() is True  # past the 5-min grace: red card
        assert screen.files["/image/status.jpg"] != amber
        assert runner.health.failures == 3
        hub.down = False
        clock.advance()
        assert await runner.tick() is True  # back to the overview
        assert runner.health.failures == 0
        await runner.aclose()

    with screen.serve() as url:
        asyncio.run(go(Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)))


def test_data_changed_card_shown_at_once(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    hub.systems["items"][1]["info"].pop("mp")
    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        assert run_ticks(runner, clock, 1) == [True]
    assert runner.health.data_changed
    assert screen.shown == "/image/status.jpg"


def test_unreachable_stops_flipping(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()
    hub.set_status("node-b", "down")

    async def go(runner):
        await runner.tick()
        await runner.flip()
        hub.down = True
        for _ in range(2):
            clock.advance()
            await runner.tick()
        assert not runner.plan.flipping
        assert screen.shown == "/image/status.jpg"
        await runner.aclose()

    with screen.serve() as url:
        asyncio.run(go(Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)))


def test_container_removed_by_beszel_also_counts(tmp_path):
    hub, screen, clock = FakeHub(), FakeScreen(), Clock()

    def world(i):
        clock.advance()
        if i == 0:
            hub.remove_container("docker-host", "app-a")

    with screen.serve() as url:
        runner = Runner(setup(tmp_path, url), now=clock, beszel_transport=hub.transport)
        run_ticks(runner, clock, 3, between=world)
    assert runner.last_status.down_containers == ("docker-host/app-a",)
