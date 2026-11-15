import asyncio
import time

import pytest

from esp_screen.device.client import DeviceError, ScreenClient, _complete, _dechunk, _parse
from tests.fakes.screen import FakeScreen


def run(fake: FakeScreen, fn):
    with fake.serve() as url:
        return asyncio.run(fn(ScreenClient(url, timeout=2)))


def test_push_uploads_then_selects_image():
    fake = FakeScreen()
    run(fake, lambda c: c.push("status.jpg", b"\xff\xd8jpeg\xff\xd9"))
    assert fake.files["/image/status.jpg"] == b"\xff\xd8jpeg\xff\xd9"
    assert fake.shown == "/image/status.jpg"
    assert fake.calls == ["POST /doUpload?dir=%2Fimage%2F", "GET /set?img=%2Fimage%2Fstatus.jpg"]


def test_tolerates_conflicting_content_length():
    """Seen on the real screen 2026-11-15: strict clients reject the upload answer."""
    fake = FakeScreen()
    run(fake, lambda c: c.upload("status.jpg", b"x" * 5000))
    assert len(fake.files["/image/status.jpg"]) == 5000


def test_same_name_replaces_file():
    fake = FakeScreen()

    async def go(c):
        await c.upload("status.jpg", b"one")
        await c.upload("status.jpg", b"two")

    run(fake, go)
    assert list(fake.files) == ["/image/status.jpg"]
    assert fake.files["/image/status.jpg"] == b"two"


def test_one_retry_on_dropped_connection():
    fake = FakeScreen(drop_next=1)
    run(fake, lambda c: c.show("status.jpg"))
    assert fake.shown == "/image/status.jpg"


def test_gives_up_after_retry():
    with pytest.raises(DeviceError, match="failed"):
        run(FakeScreen(drop_next=5), lambda c: c.show("status.jpg"))


def test_unreachable_screen_is_device_error():
    c = ScreenClient("http://127.0.0.1:9", timeout=1)  # nothing listens on the discard port
    with pytest.raises(DeviceError):
        asyncio.run(c.show("status.jpg"))


def test_http_error_is_device_error():
    with pytest.raises(DeviceError, match="HTTP 500"):
        run(FakeScreen(status_code=500), lambda c: c.upload("status.jpg", b"x"))


@pytest.mark.parametrize("name", ["../x.jpg", "status.png", "a b.jpg", "", "/image/x.jpg"])
def test_rejects_unsafe_names(name):
    with pytest.raises(ValueError):
        asyncio.run(ScreenClient("http://screen.test").upload(name, b"x"))


def test_album_settings_and_busy_answer():
    assert run(FakeScreen(autoplay=0), lambda c: c.album()).autoplay is False
    assert run(FakeScreen(autoplay=1), lambda c: c.album()).autoplay is True
    with pytest.raises(DeviceError, match="unexpected"):
        run(FakeScreen(album_busy=1), lambda c: c.album())


@pytest.mark.parametrize("url", ["", "https://screen.test", "ftp://x"])
def test_requires_http_address(url):
    with pytest.raises(ValueError):
        ScreenClient(url)


def test_address_without_scheme_and_with_port():
    c = ScreenClient("192.0.2.5:8080")
    assert (c.host, c.port) == ("192.0.2.5", 8080)


def test_parse_status_and_chunked_body():
    raw = b"HTTP/1.1 200 OK\r\nTransfer-Encoding: chunked\r\n\r\n2\r\nOK\r\n0\r\n\r\n"
    r = _parse(raw)
    assert (r.status, r.body) == (200, b"OK")
    assert _dechunk(b"zz\r\n") == b""
    with pytest.raises(DeviceError):
        _parse(b"garbage")


def test_server_that_keeps_connection_open_is_not_waited_out():
    fake = FakeScreen(keep_open=True)
    start = time.monotonic()
    run(fake, lambda c: c.push("status.jpg", b"x" * 100))
    assert fake.shown == "/image/status.jpg"
    assert time.monotonic() - start < 1.5  # finished on Content-Length, not on the 2 s timeout


def test_complete_uses_largest_content_length():
    head = b"HTTP/1.1 200 OK\r\nContent-Length: 0\r\nContent-Length: 4\r\n\r\n"
    assert not _complete(head + b"ab")
    assert _complete(head + b"abcd")
    assert _complete(b"HTTP/1.1 200 OK\r\nTransfer-Encoding: chunked\r\n\r\n0\r\n\r\n")
    assert not _complete(b"HTTP/1.1 200 OK\r\n")
