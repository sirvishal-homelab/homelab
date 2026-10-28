import pytest
from fakes.terminus import FakeTerminus

from trmnl_calendar.publish.terminus import TerminusClient, TerminusError


@pytest.fixture
def fake():
    f = FakeTerminus().start()
    yield f
    f.stop()


def client(fake, **kw):
    return TerminusClient(fake.url, kw.pop("login", "cal"), kw.pop("password", "pw"), **kw)


def test_login_and_models_by_name(fake):
    c = client(fake)
    assert c.find_model("og_plus")["id"] == 12
    with pytest.raises(TerminusError) as e:
        c.find_model("nope")
    assert e.value.kind == "config"


def test_bad_password_is_an_auth_error(fake):
    with pytest.raises(TerminusError) as e:
        client(fake, password="wrong").models()
    assert e.value.kind == "auth" and "TERMINUS_PASSWORD" in e.value.message


def test_unreachable_is_reported_plainly():
    c = TerminusClient("http://127.0.0.1:9", "a", "b", timeout=1)
    with pytest.raises(TerminusError) as e:
        c.models()
    assert e.value.kind == "unreachable"


def test_expired_token_relogs_in_once(fake):
    c = client(fake)
    c.models()
    fake.expire_tokens = True
    assert c.devices()[0]["label"] == "TRMNL"
    assert fake.calls("POST", "/login") == 2


def test_token_is_cached_then_refreshed_when_old(fake):
    now = [0.0]
    c = client(fake, clock=lambda: now[0])
    c.models()
    c.models()
    assert fake.calls("POST", "/login") == 1
    now[0] = 26 * 60
    c.models()
    assert fake.calls("POST", "/api/jwt") == 1 and fake.calls("POST", "/login") == 1


def test_create_and_patch_screen_fetches_image_during_the_call(fake, tmp_path):
    import http.server
    import threading

    png = b"\x89PNG\r\n\x1a\n" + b"x" * 20
    (tmp_path / "a.png").write_bytes(png)
    handler = lambda *a, **k: http.server.SimpleHTTPRequestHandler(  # noqa: E731
        *a, directory=str(tmp_path), **k
    )
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        uri = f"http://127.0.0.1:{srv.server_address[1]}/a.png"
        c = client(fake)
        s = c.create_screen(12, "trmnl-calendar", "Calendar", uri)
        assert fake.stored[s["id"]] == png
        c.patch_screen(s["id"], uri)
        with pytest.raises(TerminusError) as e:
            c.patch_screen(999, uri)
        assert e.value.kind == "not_found"
        fake.fail_with = 500
        with pytest.raises(TerminusError) as e:
            c.patch_screen(s["id"], uri)
        assert e.value.kind == "rejected" and e.value.status == 500
    finally:
        srv.shutdown()
        srv.server_close()
