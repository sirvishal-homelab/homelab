"""A fake SmallTV-Ultra stock firmware: a real local HTTP server with the firmware's quirks.

Quirks reproduced (docs/03): the upload answer carries two different Content-Length headers;
JSON reads right after an upload can answer a bare "OK"; connections are sometimes dropped.
"""

from __future__ import annotations

import json
import socket
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from urllib.parse import parse_qs, unquote, urlsplit


@dataclass
class FakeScreen:
    autoplay: int = 0
    drop_next: int = 0  # close this many connections without answering
    status_code: int = 200
    album_busy: int = 0  # answer /album.json with a bare "OK" this many times
    keep_open: bool = False  # ignore Connection: close and hold the socket after answering
    files: dict[str, bytes] = field(default_factory=dict)
    shown: str | None = None
    calls: list[str] = field(default_factory=list)

    # --- request handling -------------------------------------------------------------
    def _answer(self, method: str, target: str, headers: dict[str, str], body: bytes) -> bytes:
        self.calls.append(f"{method} {target}")
        if self.status_code != 200:
            return f"HTTP/1.1 {self.status_code} Error\r\nContent-Length: 0\r\n\r\n".encode()
        url = urlsplit(target)
        q = {k: v[0] for k, v in parse_qs(url.query).items()}
        if method == "POST" and url.path == "/doUpload":
            boundary = headers["content-type"].split("boundary=")[1].encode()
            part = body.split(b"--" + boundary)[1]
            head, _, data = part.partition(b"\r\n\r\n")
            name = head.split(b'filename="')[1].split(b'"')[0].decode()
            self.files[f"{q['dir']}{name}"] = data[: -len(b"\r\n")]
            page = b"<table id='list'><tr><td>status.jpg</td></tr></table>"
            # The real firmware sends two different Content-Length headers here.
            return (
                b"HTTP/1.1 200 OK\r\nContent-Type: text/html\r\nContent-Length: 0\r\n"
                b"Content-Length: " + str(len(page)).encode() + b"\r\n\r\n" + page
            )
        if method == "GET" and url.path == "/set":
            if "img" in q:
                self.shown = unquote(q["img"])
            return b"HTTP/1.1 200 OK\r\nContent-Length: 0\r\n\r\n"
        if method == "GET" and url.path == "/album.json":
            if self.album_busy > 0:
                self.album_busy -= 1
                return b"HTTP/1.1 200 OK\r\nContent-Length: 2\r\n\r\nOK"
            page = json.dumps({"autoplay": self.autoplay, "i_i": 5}).encode()
            return b"HTTP/1.1 200 OK\r\nContent-Length: %d\r\n\r\n%s" % (len(page), page)
        return b"HTTP/1.1 404 Not Found\r\nContent-Length: 3\r\n\r\n404"

    def _handle(self, conn: socket.socket) -> None:
        with conn:
            conn.settimeout(5)
            if self.drop_next > 0:
                self.drop_next -= 1
                return
            buf = b""
            while b"\r\n\r\n" not in buf:
                chunk = conn.recv(65536)
                if not chunk:
                    return
                buf += chunk
            head, _, body = buf.partition(b"\r\n\r\n")
            lines = head.decode("latin-1").split("\r\n")
            method, target, _ = lines[0].split(" ", 2)
            headers = {
                k.strip().lower(): v.strip() for k, v in (ln.split(":", 1) for ln in lines[1:])
            }
            need = int(headers.get("content-length", "0"))
            while len(body) < need:
                chunk = conn.recv(65536)
                if not chunk:
                    break
                body += chunk
            conn.sendall(self._answer(method, target, headers, body))
            if self.keep_open:
                try:
                    conn.recv(1)  # hold until the client hangs up
                except OSError:
                    pass

    # --- server -----------------------------------------------------------------------
    @contextmanager
    def serve(self):
        """Run on 127.0.0.1 in a background thread; yields the base URL."""
        srv = socket.create_server(("127.0.0.1", 0))
        srv.settimeout(0.2)
        stop = threading.Event()

        def loop():
            while not stop.is_set():
                try:
                    conn, _ = srv.accept()
                except TimeoutError:
                    continue
                except OSError:
                    return
                self._handle(conn)

        t = threading.Thread(target=loop, daemon=True)
        t.start()
        try:
            yield f"http://127.0.0.1:{srv.getsockname()[1]}"
        finally:
            stop.set()
            t.join(2)
            srv.close()
