import asyncio

import pytest

from esp_screen.device.client import DeviceError, ScreenClient
from tests.fakes.screen import FakeScreen


def run(coro):
    return asyncio.run(coro)


def client(fake: FakeScreen) -> ScreenClient:
    return ScreenClient("http://screen.test", transport=fake.transport)


def test_push_uploads_then_selects_image():
    fake = FakeScreen()

    async def go():
        async with client(fake) as c:
            await c.push("status.jpg", b"\xff\xd8jpeg\xff\xd9")

    run(go())
    assert fake.files["/image/status.jpg"] == b"\xff\xd8jpeg\xff\xd9"
    assert fake.shown == "/image/status.jpg"
    assert fake.calls == ["POST /doUpload?dir=%2Fimage%2F", "GET /set?img=%2Fimage%2Fstatus.jpg"]


def test_same_name_replaces_file():
    fake = FakeScreen()

    async def go():
        async with client(fake) as c:
            await c.upload("status.jpg", b"one")
            await c.upload("status.jpg", b"two")

    run(go())
    assert list(fake.files) == ["/image/status.jpg"]
    assert fake.files["/image/status.jpg"] == b"two"


def test_one_retry_on_dropped_connection():
    fake = FakeScreen(fail_next=1)

    async def go():
        async with client(fake) as c:
            await c.show("status.jpg")

    run(go())
    assert fake.shown == "/image/status.jpg"


def test_gives_up_after_retry():
    fake = FakeScreen(fail_next=5)

    async def go():
        async with client(fake) as c:
            await c.show("status.jpg")

    with pytest.raises(DeviceError, match="failed"):
        run(go())


def test_http_error_is_device_error():
    fake = FakeScreen(status_code=500)

    async def go():
        async with client(fake) as c:
            await c.upload("status.jpg", b"x")

    with pytest.raises(DeviceError, match="HTTP 500"):
        run(go())


@pytest.mark.parametrize("name", ["../x.jpg", "status.png", "a b.jpg", "", "/image/x.jpg"])
def test_rejects_unsafe_names(name):
    async def go():
        async with client(FakeScreen()) as c:
            await c.upload(name, b"x")

    with pytest.raises(ValueError):
        run(go())


def test_album_settings():
    async def go(fake):
        async with client(fake) as c:
            return await c.album()

    assert run(go(FakeScreen(autoplay=0))).autoplay is False
    assert run(go(FakeScreen(autoplay=1))).autoplay is True


def test_requires_address():
    with pytest.raises(ValueError):
        ScreenClient("")
