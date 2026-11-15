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
