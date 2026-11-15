"""A fake Beszel hub (PocketBase API) serving the recorded 0.21.2 fixtures (docs/08)."""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from pathlib import Path

import httpx

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "beszel-0.21.2"
USER, PASSWORD = "esp-screen@example.com", "pw"


def load(name: str) -> dict:
    return json.loads((FIXTURES / f"{name}.json").read_text())


@dataclass
class FakeHub:
    systems: dict = field(default_factory=lambda: load("systems"))
    containers: dict = field(default_factory=lambda: load("containers"))
    expire_next: int = 0  # answer 401 this many times to authorised reads (expired token)
    down: bool = False  # refuse connections
    logins: int = 0
    tokens: set[str] = field(default_factory=set)

    def handler(self, request: httpx.Request) -> httpx.Response:
        if self.down:
            raise httpx.ConnectError("refused", request=request)
        path = request.url.path
        if request.method == "POST" and path == "/api/collections/users/auth-with-password":
            body = json.loads(request.content)
            if (body.get("identity"), body.get("password")) != (USER, PASSWORD):
                return httpx.Response(400, json={"message": "Failed to authenticate."})
            self.logins += 1
            token = f"tok{self.logins}"
            self.tokens.add(token)
            return httpx.Response(200, json={"token": token, "record": {"role": "readonly"}})
        if request.headers.get("authorization") not in self.tokens:
            return httpx.Response(401, json={"message": "unauthorised"})
        if self.expire_next > 0:
            self.expire_next -= 1
            self.tokens.clear()
            return httpx.Response(401, json={"message": "token expired"})
        if path == "/api/collections/systems/records":
            return httpx.Response(200, json=copy.deepcopy(self.systems))
        if path == "/api/collections/containers/records":
            return httpx.Response(200, json=copy.deepcopy(self.containers))
        return httpx.Response(404, json={"message": "not found"})

    @property
    def transport(self) -> httpx.MockTransport:
        return httpx.MockTransport(self.handler)

    # helpers to change the world between reads
    def set_status(self, name: str, status: str) -> None:
        for s in self.systems["items"]:
            if s["name"] == name:
                s["status"] = status

    def stop_container(self, system_name: str, name: str) -> None:
        sid = next(s["id"] for s in self.systems["items"] if s["name"] == system_name)
        self.containers["items"] = [
            c for c in self.containers["items"] if not (c["system"] == sid and c["name"] == name)
        ]
