"""HTTP client for the GeekMagic SmallTV-Ultra stock firmware (docs/03, tested on Ultra-V9.0.50).

The firmware API is undocumented: keep every request it makes in this one module.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from urllib.parse import quote

import httpx

IMAGE_DIR = "/image/"
_SAFE_NAME = re.compile(r"^[a-z0-9][a-z0-9._-]{0,40}\.(jpg|gif)$")


class DeviceError(Exception):
    """The screen could not be reached or refused a request."""


@dataclass(frozen=True)
class AlbumSettings:
    autoplay: bool
    interval_s: int


class ScreenClient:
    def __init__(
        self,
        base_url: str,
        *,
        timeout: float = 5.0,
        transport: httpx.AsyncBaseTransport | None = None,
    ) -> None:
        if not base_url:
            raise ValueError("screen address is not set")
        self._client = httpx.AsyncClient(
            base_url=base_url.rstrip("/"), timeout=timeout, transport=transport
        )

    async def aclose(self) -> None:
        await self._client.aclose()

    async def __aenter__(self) -> ScreenClient:
        return self

    async def __aexit__(self, *exc) -> None:
        await self.aclose()

    async def _request(self, method: str, url: str, **kw) -> httpx.Response:
        last: Exception | None = None
        for _ in range(2):  # one retry: the ESP8266 sometimes drops a connection
            try:
                r = await self._client.request(method, url, **kw)
            except httpx.TransportError as exc:
                last = exc
                continue
            if r.status_code >= 400:
                raise DeviceError(f"{method} {url} returned HTTP {r.status_code}")
            return r
        raise DeviceError(f"{method} {url} failed: {type(last).__name__}: {last}")

    async def upload(self, name: str, data: bytes) -> None:
        """Store an image in /image/ (an existing file with the same name is replaced)."""
        if not _SAFE_NAME.match(name):
            raise ValueError(f"unsafe image name: {name!r}")
        mime = "image/gif" if name.endswith(".gif") else "image/jpeg"
        await self._request(
            "POST",
            "/doUpload",
            params={"dir": IMAGE_DIR},
            files={"file": (name, data, mime)},
        )

    async def show(self, name: str) -> None:
        """Select the image to display (visible on the Photo Album theme)."""
        await self._request("GET", f"/set?img={quote(IMAGE_DIR + name, safe='')}")

    async def push(self, name: str, data: bytes) -> None:
        await self.upload(name, data)
        await self.show(name)

    async def album(self) -> AlbumSettings:
        r = await self._request("GET", "/album.json")
        try:
            body = r.json()
            return AlbumSettings(bool(int(body.get("autoplay", 0))), int(body.get("i_i", 0)))
        except (ValueError, TypeError, AttributeError) as exc:
            raise DeviceError(f"unexpected /album.json: {r.text[:100]!r}") from exc
