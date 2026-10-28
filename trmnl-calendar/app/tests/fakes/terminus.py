"""A small fake Terminus for contract tests: same paths, shapes and behaviour as the spike saw.

It fetches the screen's URI *during* POST/PATCH, like the real one, and keeps the bytes.
"""

from __future__ import annotations

import json
import threading
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

PNG = b"\x89PNG\r\n\x1a\n"


def _model(mid: int, name: str, label: str, bit_depth: int) -> dict:
    return {
        "id": mid,
        "name": name,
        "label": label,
        "width": 800,
        "height": 480,
        "bit_depth": bit_depth,
        "mime_type": "image/png",
    }


class FakeTerminus:
    def __init__(self, login="cal", password="pw", model_name="og_plus", bit_depth=2):
        self.login, self.password = login, password
        self.token_n = 0
        self.valid_tokens: set[str] = set()
        self.models = [
            _model(12, model_name, "TRMNL OG", bit_depth),
            _model(2, "og_png", "TRMNL OG 1-bit", 1),
        ]
        self.devices = [
            {
                "id": 7,
                "model_id": 12,
                "playlist_id": 7,
                "label": "TRMNL",
                "firmware_version": "1.8.6",
                "battery_voltage": 4.2,
                "wifi_signal": -26,
                "refresh_rate": 900,
                "synced_at": "2026-10-28T13:00:00+00:00",
            }
        ]
        self.playlists = [{"id": 7, "name": "device_7", "items": []}]
        self.screens: dict[int, dict] = {}
        self.stored: dict[int, bytes] = {}
        self.next_id = 20
        self.requests: list[tuple[str, str]] = []
        self.fail_with: int | None = None  # make the next screen call fail with this status
        self.expire_tokens = False
        fake = self

        class H(BaseHTTPRequestHandler):
            def log_message(self, *a):  # silence
                pass

            def _send(self, status, obj):
                body = json.dumps(obj).encode()
                self.send_response(status)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def _body(self):
                n = int(self.headers.get("Content-Length") or 0)
                return json.loads(self.rfile.read(n) or b"{}")

            def _auth(self):
                return self.headers.get("Authorization") in fake.valid_tokens

            def handle_any(self, method):
                path = self.path
                fake.requests.append((method, path))
                if method == "POST" and path == "/login":
                    b = self._body()
                    if b.get("login") != fake.login or b.get("password") != fake.password:
                        return self._send(401, {"error": "Invalid"})
                    fake.token_n += 1
                    tok = f"tok{fake.token_n}"
                    fake.valid_tokens.add(tok)
                    return self._send(200, {"access_token": tok, "refresh_token": f"r{tok}"})
                if method == "POST" and path == "/api/jwt":
                    if not self._auth():
                        return self._send(401, {})
                    fake.token_n += 1
                    tok = f"tok{fake.token_n}"
                    fake.valid_tokens.add(tok)
                    return self._send(200, {"access_token": tok, "refresh_token": f"r{tok}"})
                if not self._auth():
                    return self._send(401, {"error": "Please log in"})
                if fake.expire_tokens:
                    fake.valid_tokens.clear()
                    fake.expire_tokens = False
                    return self._send(401, {"error": "expired"})
                if method == "GET" and path == "/api/models":
                    return self._send(200, {"data": fake.models})
                if method == "GET" and path == "/api/devices":
                    return self._send(200, {"data": fake.devices})
                if method == "GET" and path == "/api/playlists":
                    return self._send(200, {"data": fake.playlists})
                if method == "GET" and path == "/api/screens":
                    return self._send(200, {"data": list(fake.screens.values())})
                if path.startswith("/api/screens"):
                    return self.screens_route(method, path)
                return self._send(404, {"error": "nope"})

            def screens_route(self, method, path):
                parts = path.strip("/").split("/")
                sid = int(parts[2]) if len(parts) > 2 else None
                if method == "GET" and sid in fake.screens:
                    return self._send(200, {"data": fake.screens[sid]})
                if method in ("POST", "PATCH"):
                    if fake.fail_with:
                        status, fake.fail_with = fake.fail_with, None
                        return self._send(status, {"error": "boom"})
                    s = self._body().get("screen", {})
                    if method == "PATCH" and sid not in fake.screens:
                        return self._send(404, {"error": "Not found"})
                    try:
                        with urllib.request.urlopen(s["uri"], timeout=5) as r:  # noqa: S310
                            data = r.read()
                    except Exception as exc:  # the real one reports a fetch failure as 422
                        return self._send(422, {"error": f"could not fetch image: {exc}"})
                    if not data.startswith(PNG):
                        return self._send(422, {"error": "not a PNG"})
                    if method == "POST":
                        sid = fake.next_id
                        fake.next_id += 1
                        fake.screens[sid] = {
                            "id": sid,
                            "model_id": int(s["model_id"]),
                            "name": s["name"],
                            "label": s["label"],
                        }
                    fake.screens[sid]["uri"] = s["uri"]
                    fake.stored[sid] = data
                    return self._send(200, {"data": fake.screens[sid]})
                return self._send(404, {"error": "nope"})

            def do_GET(self):
                self.handle_any("GET")

            def do_POST(self):
                self.handle_any("POST")

            def do_PATCH(self):
                self.handle_any("PATCH")

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), H)
        self.url = f"http://127.0.0.1:{self.server.server_address[1]}"
        self._thread = threading.Thread(target=self.server.serve_forever, daemon=True)

    def start(self):
        self._thread.start()
        return self

    def stop(self):
        self.server.shutdown()
        self.server.server_close()

    def calls(self, method: str, prefix: str) -> int:
        return sum(1 for m, p in self.requests if m == method and p.startswith(prefix))
