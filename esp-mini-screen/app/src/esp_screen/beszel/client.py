"""Beszel hub client (PocketBase REST API), tested against Beszel 0.21.2 (docs/08).

Beszel's API is PocketBase internals and may change in minor releases (R-1). All knowledge of
its JSON lives in `to_system` and `to_container`; anything unexpected raises BeszelDataError so
the screen can say "Beszel data changed" instead of showing wrong numbers (F-16).
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime, timedelta
from typing import Any

import httpx

from esp_screen.domain.model import Container, Snapshot, System, SystemState

# A running container's record is rewritten with every agent report, with the same `updated`
# time as its system. A stopped container's record stays behind with an old time (docs/08 B-2,
# corrected 2026-11-15), so "running" means "updated in the system's latest report".
FRESH_WITHIN = timedelta(seconds=15)


class BeszelError(Exception):
    """The hub could not be reached, refused the login, or answered with an HTTP error."""


class BeszelDataError(BeszelError):
    """The hub answered, but not in the shape this app was built for."""


def _need(rec: dict, key: str, kind: type | tuple[type, ...], where: str) -> Any:
    if key not in rec:
        raise BeszelDataError(f"{where}: field '{key}' is missing")
    value = rec[key]
    if not isinstance(value, kind):
        raise BeszelDataError(f"{where}: field '{key}' has type {type(value).__name__}")
    return value


def _system_updated(rec: dict, where: str) -> datetime:
    raw = _need(rec, "updated", str, where)
    try:
        return datetime.fromisoformat(raw.replace(" ", "T").replace("Z", "+00:00"))
    except ValueError as exc:
        raise BeszelDataError(f"{where}: field 'updated' is not a time: {raw!r}") from exc


def _container_updated(rec: dict, where: str) -> datetime:
    ms = _need(rec, "updated", (int, float), where)
    return datetime.fromtimestamp(ms / 1000, tz=UTC)


def running_containers(systems: list[dict], containers: list[dict]) -> list[Container]:
    """Containers whose record came with their system's latest report; stale records dropped."""
    last_report = {
        _need(s, "id", str, "systems"): _system_updated(s, f"systems/{s.get('id', '?')}")
        for s in systems
    }
    out = []
    for rec in containers:
        c = to_container(rec)
        seen = _container_updated(rec, f"containers/{rec.get('id', '?')}")
        report = last_report.get(c.system_id)
        if report is not None and seen >= report - FRESH_WITHIN:
            out.append(c)
    return list(dict.fromkeys(out))  # a recreated container can leave an old twin behind


def to_system(rec: dict) -> System:
    where = f"systems/{rec.get('id', '?')}"
    status = _need(rec, "status", str, where)
    try:
        state = SystemState(status)
    except ValueError as exc:
        raise BeszelDataError(f"{where}: unknown status '{status}'") from exc
    info = _need(rec, "info", dict, where)
    number = (int, float)
    return System(
        id=_need(rec, "id", str, where),
        name=_need(rec, "name", str, where),
        state=state,
        cpu_pct=float(_need(info, "cpu", number, where + ".info")) if info else 0.0,
        mem_pct=float(_need(info, "mp", number, where + ".info")) if info else 0.0,
    )


def to_container(rec: dict) -> Container:
    where = f"containers/{rec.get('id', '?')}"
    return Container(
        system_id=_need(rec, "system", str, where), name=_need(rec, "name", str, where)
    )


class BeszelClient:
    def __init__(
        self,
        base_url: str,
        user: str,
        password: str,
        *,
        timeout: float = 10.0,
        transport: httpx.AsyncBaseTransport | None = None,
        now: Callable[[], datetime] = lambda: datetime.now(UTC),
    ) -> None:
        self._client = httpx.AsyncClient(base_url=base_url, timeout=timeout, transport=transport)
        self._user = user
        self._password = password
        self._token: str | None = None
        self._now = now

    async def aclose(self) -> None:
        await self._client.aclose()

    async def _login(self) -> None:
        try:
            r = await self._client.post(
                "/api/collections/users/auth-with-password",
                json={"identity": self._user, "password": self._password},
            )
        except httpx.HTTPError as exc:
            raise BeszelError(f"cannot reach Beszel: {type(exc).__name__}") from exc
        if r.status_code in (400, 401, 403):
            raise BeszelError("Beszel refused the login (check BESZEL_USER / BESZEL_PASSWORD)")
        if r.status_code >= 400:
            raise BeszelError(f"Beszel login returned HTTP {r.status_code}")
        token = (
            r.json().get("token")
            if r.headers.get("content-type", "").startswith("application/json")
            else None
        )
        if not isinstance(token, str) or not token:
            raise BeszelDataError("login answer has no token")
        self._token = token

    async def _items(self, collection: str, per_page: int) -> list[dict]:
        for attempt in range(2):
            if self._token is None:
                await self._login()
            try:
                r = await self._client.get(
                    f"/api/collections/{collection}/records",
                    params={"perPage": per_page, "skipTotal": 1},
                    headers={"Authorization": self._token or ""},
                )
            except httpx.HTTPError as exc:
                raise BeszelError(f"cannot reach Beszel: {type(exc).__name__}") from exc
            if r.status_code in (401, 403) and attempt == 0:
                self._token = None  # expired token: log in again once
                continue
            if r.status_code >= 400:
                raise BeszelError(f"{collection}: HTTP {r.status_code}")
            try:
                body = r.json()
            except ValueError as exc:
                raise BeszelDataError(f"{collection}: answer is not JSON") from exc
            items = body.get("items") if isinstance(body, dict) else None
            if not isinstance(items, list):
                raise BeszelDataError(f"{collection}: answer has no 'items' list")
            return items
        raise BeszelError(f"{collection}: not authorised after a fresh login")

    async def snapshot(self) -> Snapshot:
        system_recs = await self._items("systems", 500)
        systems = [to_system(r) for r in system_recs]
        containers = running_containers(system_recs, await self._items("containers", 1000))
        return Snapshot(tuple(systems), tuple(containers), self._now())
