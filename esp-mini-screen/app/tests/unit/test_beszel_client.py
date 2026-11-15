import asyncio

import pytest

from esp_screen.beszel.client import BeszelClient, BeszelDataError, BeszelError
from esp_screen.domain.model import SystemState
from tests.fakes.beszel import PASSWORD, USER, FakeHub


def snap(hub: FakeHub, user=USER, password=PASSWORD):
    async def go():
        c = BeszelClient("http://beszel:8090", user, password, transport=hub.transport)
        try:
            return await c.snapshot()
        finally:
            await c.aclose()

    return asyncio.run(go())


def test_reads_recorded_hub():
    s = snap(FakeHub())
    by_name = {x.name: x for x in s.systems}
    assert len(s.systems) == 7
    assert by_name["node-a"].state is SystemState.UP
    assert by_name["node-a"].cpu_pct == pytest.approx(0.91)
    assert by_name["node-a"].mem_pct == pytest.approx(6.97)
    assert by_name["backup"].state is SystemState.PAUSED
    assert len(s.containers) == 9
    assert {c.name for c in s.containers} >= {"beszel-agent", "watchtower", "app-a"}


def test_wrong_password_is_clear():
    with pytest.raises(BeszelError, match="refused the login"):
        snap(FakeHub(), password="nope")


def test_expired_token_logs_in_again_once():
    hub = FakeHub()

    async def go():
        c = BeszelClient("http://beszel:8090", USER, PASSWORD, transport=hub.transport)
        await c.snapshot()
        hub.expire_next = 1
        await c.snapshot()
        await c.aclose()

    asyncio.run(go())
    assert hub.logins == 2


def test_token_is_reused_between_reads():
    hub = FakeHub()

    async def go():
        c = BeszelClient("http://beszel:8090", USER, PASSWORD, transport=hub.transport)
        for _ in range(3):
            await c.snapshot()
        await c.aclose()

    asyncio.run(go())
    assert hub.logins == 1


def test_unreachable_hub():
    with pytest.raises(BeszelError, match="cannot reach"):
        snap(FakeHub(down=True))


@pytest.mark.parametrize(
    "mutate, message",
    [
        (lambda s: s.pop("status"), "'status' is missing"),
        (lambda s: s.__setitem__("status", "sleeping"), "unknown status"),
        (lambda s: s["info"].pop("mp"), "'mp' is missing"),
        (lambda s: s["info"].__setitem__("cpu", "high"), "'cpu' has type str"),
    ],
)
def test_changed_data_shape_is_data_error(mutate, message):
    hub = FakeHub()
    mutate(hub.systems["items"][1])
    with pytest.raises(BeszelDataError, match=message):
        snap(hub)


def test_missing_items_list_is_data_error():
    hub = FakeHub()
    hub.containers = {"records": []}
    with pytest.raises(BeszelDataError, match="no 'items'"):
        snap(hub)


def test_stopped_container_record_is_not_running():
    """Seen on the real hub 2026-11-15: a stopped container's record stays with an old time."""
    hub = FakeHub()
    hub.stop_container("docker-host", "watchtower")
    names = {(c.system_id, c.name) for c in snap(hub).containers}
    assert ("grh0kefzwq2t2g3", "watchtower") not in names
    assert ("grh0kefzwq2t2g3", "app-a") in names
    assert len(names) == 8


def test_old_twin_of_recreated_container_is_ignored():
    hub = FakeHub()
    twin = dict(hub.containers["items"][1], id="oldid", updated=1794738213350 - 600_000)
    hub.containers["items"].append(twin)
    assert len(snap(hub).containers) == 9


def test_containers_of_unknown_system_are_ignored():
    hub = FakeHub()
    hub.containers["items"].append(dict(hub.containers["items"][0], system="nosuch"))
    assert len(snap(hub).containers) == 9


def test_bad_update_times_are_data_errors():
    hub = FakeHub()
    hub.systems["items"][0]["updated"] = "yesterday"
    with pytest.raises(BeszelDataError, match="not a time"):
        snap(hub)
    hub = FakeHub()
    hub.containers["items"][0]["updated"] = "soon"
    with pytest.raises(BeszelDataError, match="'updated' has type str"):
        snap(hub)
