"""HTTP client for the GeekMagic SmallTV-Ultra stock firmware (docs/03, tested on Ultra-V9.0.50).

The firmware's web server is not strict HTTP: its upload answer carries two different
`Content-Length` headers (seen 2026-11-15), which strict clients such as httpx/h11 reject.
So this module speaks plain HTTP/1.1 over a socket with `Connection: close`, reads until the
server closes, and only trusts the status line. Every request the firmware gets is in here.
"""

from __future__ import annotations

import asyncio
import json
import re
import uuid
from dataclasses import dataclass
from urllib.parse import quote, urlsplit

IMAGE_DIR = "/image/"
MAX_RESPONSE = 1_000_000
_SAFE_NAME = re.compile(r"^[a-z0-9][a-z0-9._-]{0,40}\.(jpg|gif)$")


class DeviceError(Exception):
    """The screen could not be reached or refused a request."""


@dataclass(frozen=True)
class AlbumSettings:
    autoplay: bool
    interval_s: int


@dataclass(frozen=True)
class Reply:
    status: int
    body: bytes

    @property
    def text(self) -> str:
        return self.body.decode("utf-8", "replace")


def _parse(raw: bytes) -> Reply:
    head, sep, body = raw.partition(b"\r\n\r\n")
    if not sep:
        raise DeviceError("incomplete HTTP response")
    lines = head.split(b"\r\n")
    m = re.match(rb"HTTP/1\.[01] (\d{3})", lines[0])
    if not m:
        raise DeviceError(f"not an HTTP response: {lines[0][:40]!r}")
    headers = [line.split(b":", 1)[0].strip().lower() for line in lines[1:]]
    if b"transfer-encoding" in headers and b"chunked" in head.lower():
        body = _dechunk(body)
    return Reply(int(m.group(1)), body)


def _dechunk(data: bytes) -> bytes:
    out, rest = bytearray(), data
    while rest:
        size_line, _, rest = rest.partition(b"\r\n")
        try:
            size = int(size_line.split(b";")[0].strip() or b"0", 16)
        except ValueError:
            break  # tolerate a malformed tail; the status line is what counts
        if size == 0:
            break
        out += rest[:size]
        rest = rest[size + 2 :]
    return bytes(out)


def _complete(raw: bytes) -> bool:
    head, sep, body = raw.partition(b"\r\n\r\n")
    if not sep:
        return False
    low = head.lower()
    if b"transfer-encoding:" in low and b"chunked" in low:
        return body.endswith(b"0\r\n\r\n")
    lengths = [int(n) for n in re.findall(rb"\r\ncontent-length:\s*(\d+)", low)]
    return bool(lengths) and len(body) >= max(lengths)


class ScreenClient:
    def __init__(self, base_url: str, *, timeout: float = 5.0) -> None:
        if not base_url:
            raise ValueError("screen address is not set")
        parts = urlsplit(base_url if "://" in base_url else f"http://{base_url}")
        if parts.scheme != "http" or not parts.hostname:
            raise ValueError(f"screen address must be http://host[:port], got {base_url!r}")
        self.host = parts.hostname
        self.port = parts.port or 80
        self.timeout = timeout

    async def __aenter__(self) -> ScreenClient:
        return self

    async def __aexit__(self, *exc) -> None:
        return None

    async def aclose(self) -> None:  # kept for symmetry; there is no pooled connection
        return None

    async def _once(self, method: str, target: str, body: bytes, headers: dict[str, str]) -> Reply:
        reader, writer = await asyncio.wait_for(
            asyncio.open_connection(self.host, self.port), self.timeout
        )
        try:
            lines = [
                f"{method} {target} HTTP/1.1",
                f"Host: {self.host}" + ("" if self.port == 80 else f":{self.port}"),
                "Connection: close",
                "User-Agent: esp-mini-screen",
                f"Content-Length: {len(body)}",
                *(f"{k}: {v}" for k, v in headers.items()),
            ]
            writer.write(("\r\n".join(lines) + "\r\n\r\n").encode("latin-1") + body)
            await writer.drain()
            raw = await self._read_reply(reader)
        finally:
            writer.close()
            try:
                await writer.wait_closed()
            except OSError:
                pass
        return _parse(raw)

    async def _read_reply(self, reader: asyncio.StreamReader) -> bytes:
        """Read until the reply is complete, the server closes, or it goes quiet.

        Complete means: all body bytes announced (the largest Content-Length, since the firmware
        may send two), or the end of a chunked body. A server that keeps the connection open is
        tolerated once the headers have arrived.
        """
        raw = b""
        while len(raw) < MAX_RESPONSE:
            try:
                more = await asyncio.wait_for(reader.read(65536), self.timeout)
            except TimeoutError:
                if b"\r\n\r\n" in raw:
                    break
                raise
            if not more:
                break
            raw += more
            if _complete(raw):
                break
        return raw

    async def _request(
        self, method: str, target: str, body: bytes = b"", headers: dict[str, str] | None = None
    ) -> Reply:
        last: Exception | None = None
        for _ in range(2):  # one retry: the ESP8266 sometimes drops a connection
            try:
                reply = await self._once(method, target, body, headers or {})
            except (OSError, TimeoutError, DeviceError) as exc:
                last = exc
                continue
            if reply.status >= 400:
                raise DeviceError(f"{method} {target} returned HTTP {reply.status}")
            return reply
        raise DeviceError(f"{method} {target} failed: {type(last).__name__}: {last}")

    async def upload(self, name: str, data: bytes) -> None:
        """Store an image in /image/ (an existing file with the same name is replaced)."""
        if not _SAFE_NAME.match(name):
            raise ValueError(f"unsafe image name: {name!r}")
        mime = "image/gif" if name.endswith(".gif") else "image/jpeg"
        boundary = f"----esp{uuid.uuid4().hex}"
        head = (
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="file"; filename="{name}"\r\n'
            f"Content-Type: {mime}\r\n\r\n"
        )
        body = head.encode() + data + f"\r\n--{boundary}--\r\n".encode()
        await self._request(
            "POST",
            f"/doUpload?dir={quote(IMAGE_DIR, safe='')}",
            body,
            {"Content-Type": f"multipart/form-data; boundary={boundary}"},
        )

    async def show(self, name: str) -> None:
        """Select the image to display (visible on the Photo Album theme)."""
        await self._request("GET", f"/set?img={quote(IMAGE_DIR + name, safe='')}")

    async def push(self, name: str, data: bytes) -> None:
        await self.upload(name, data)
        await self.show(name)

    async def album(self) -> AlbumSettings:
        reply = await self._request("GET", "/album.json")
        try:
            body = json.loads(reply.body)
            return AlbumSettings(bool(int(body.get("autoplay", 0))), int(body.get("i_i", 0)))
        except (ValueError, TypeError, AttributeError) as exc:
            raise DeviceError(f"unexpected /album.json: {reply.text[:100]!r}") from exc
