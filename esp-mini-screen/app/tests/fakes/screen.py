"""A fake SmallTV-Ultra stock firmware, as an httpx MockTransport (docs/03 behaviour)."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from urllib.parse import unquote

import httpx


@dataclass
class FakeScreen:
    autoplay: int = 0
    fail_next: int = 0  # number of connections to drop before answering
    status_code: int = 200
    files: dict[str, bytes] = field(default_factory=dict)
    shown: str | None = None
    calls: list[str] = field(default_factory=list)

    def handler(self, request: httpx.Request) -> httpx.Response:
        self.calls.append(f"{request.method} {request.url.raw_path.decode()}")
        if self.fail_next > 0:
            self.fail_next -= 1
            raise httpx.ConnectError("connection dropped", request=request)
        if self.status_code != 200:
            return httpx.Response(self.status_code)
        path = request.url.path
        if request.method == "POST" and path == "/doUpload":
            directory = request.url.params.get("dir")
            body = request.content
            ctype = request.headers["content-type"]
            boundary = ctype.split("boundary=")[1].encode()
            part = body.split(b"--" + boundary)[1]
            head, _, data = part.partition(b"\r\n\r\n")
            name = head.split(b'filename="')[1].split(b'"')[0].decode()
            self.files[f"{directory}{name}"] = data.rstrip(b"\r\n")
            return httpx.Response(200, text="<table id='list'></table>")
        if request.method == "GET" and path == "/set":
            img = request.url.params.get("img")
            if img is not None:
                self.shown = unquote(img)
            return httpx.Response(200, text="")
        if request.method == "GET" and path == "/album.json":
            return httpx.Response(200, text=json.dumps({"autoplay": self.autoplay, "i_i": 5}))
        return httpx.Response(404, text="404")

    @property
    def transport(self) -> httpx.MockTransport:
        return httpx.MockTransport(self.handler)
