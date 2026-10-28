"""The only module that knows Terminus's API (docs/05, verified by the spike in docs/14).

Standard library only (urllib), so it adds no dependency. The client is synchronous: the scheduler
runs it in a worker thread. A 200 from PATCH or POST means Terminus has already fetched and stored
the image (it fetches the URI during the call).
"""

from __future__ import annotations

import json
import ssl
import time
import urllib.error
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass

TOKEN_LIFETIME = 25 * 60  # access tokens last about 30 minutes; refresh a little early


class TerminusError(Exception):
    """kind: unreachable | auth | not_found | rejected | bad_response | config."""

    def __init__(self, kind: str, message: str, status: int | None = None, body: str = ""):
        super().__init__(message)
        self.kind = kind
        self.message = message
        self.status = status
        self.body = body


@dataclass(frozen=True)
class Reply:
    status: int
    body: bytes


Transport = Callable[[str, str, dict, bytes | None, float], Reply]


def urllib_transport(verify_tls: bool = True) -> Transport:
    ctx = None if verify_tls else ssl._create_unverified_context()  # noqa: S323 (owner opt-out)

    def send(method: str, url: str, headers: dict, body: bytes | None, timeout: float) -> Reply:
        req = urllib.request.Request(url, data=body, method=method, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:  # noqa: S310
                return Reply(r.status, r.read())
        except urllib.error.HTTPError as exc:
            return Reply(exc.code, exc.read())
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            reason = getattr(exc, "reason", exc)
            raise TerminusError("unreachable", f"Could not reach Terminus: {reason}") from exc

    return send


class TerminusClient:
    def __init__(
        self,
        base_url: str,
        login: str,
        password: str,
        *,
        transport: Transport | None = None,
        verify_tls: bool = True,
        timeout: float = 20.0,
        clock: Callable[[], float] = time.monotonic,
    ):
        self.base = base_url.rstrip("/")
        self._login = login
        self._password = password
        self._send = transport or urllib_transport(verify_tls)
        self._timeout = timeout
        self._clock = clock
        self._access: str | None = None
        self._refresh: str | None = None
        self._issued = 0.0

    # ---- plumbing

    def _raw(self, method: str, path: str, payload: dict | None, token: str | None) -> Reply:
        headers = {"Accept": "application/json"}
        body = None
        if payload is not None:
            headers["Content-Type"] = "application/json"
            body = json.dumps(payload).encode()
        if token:
            headers["Authorization"] = token  # bare token, verified in the spike
        return self._send(method, self.base + path, headers, body, self._timeout)

    @staticmethod
    def _json(reply: Reply):
        try:
            return json.loads(reply.body.decode("utf-8") or "{}")
        except ValueError as exc:
            raise TerminusError(
                "bad_response", "Terminus sent a reply this app could not read.", reply.status
            ) from exc

    @staticmethod
    def _text(reply: Reply) -> str:
        return reply.body.decode("utf-8", "replace")[:300]

    def login(self) -> None:
        reply = self._raw(
            "POST", "/login", {"login": self._login, "password": self._password}, None
        )
        if reply.status in (401, 403, 422):
            raise TerminusError(
                "auth",
                "Terminus did not accept the login. Check TERMINUS_LOGIN and TERMINUS_PASSWORD.",
                reply.status,
                self._text(reply),
            )
        if reply.status != 200:
            raise TerminusError(
                "bad_response", f"Terminus login returned HTTP {reply.status}.", reply.status
            )
        data = self._json(reply)
        token = data.get("access_token")
        if not token:
            raise TerminusError("bad_response", "Terminus login did not return a token.", 200)
        self._access, self._refresh, self._issued = token, data.get("refresh_token"), self._clock()

    def _refresh_token(self) -> bool:
        if not self._refresh:
            return False
        try:
            reply = self._raw("POST", "/api/jwt", {"refresh_token": self._refresh}, self._access)
        except TerminusError:
            return False
        if reply.status != 200:
            return False
        data = self._json(reply)
        if not data.get("access_token"):
            return False
        self._access = data["access_token"]
        self._refresh = data.get("refresh_token", self._refresh)
        self._issued = self._clock()
        return True

    def _token(self) -> str:
        if self._access is None:
            self.login()
        elif self._clock() - self._issued > TOKEN_LIFETIME and not self._refresh_token():
            self.login()
        return self._access or ""

    def call(self, method: str, path: str, payload: dict | None = None):
        """Authenticated request; one re-login on 401. Returns the decoded `data` field."""
        reply = self._raw(method, path, payload, self._token())
        if reply.status == 401:
            self._access = None
            reply = self._raw(method, path, payload, self._token())
        # Terminus answers 422 "Unable to find screen: N" (not 404) for a screen that was deleted
        gone = reply.status == 422 and "unable to find" in self._text(reply).lower()
        if reply.status == 404 or gone:
            raise TerminusError("not_found", f"Terminus has no {path}.", 404, self._text(reply))
        if reply.status == 401:
            raise TerminusError("auth", "Terminus keeps refusing the login.", 401)
        if reply.status >= 400:
            raise TerminusError(
                "rejected",
                f"Terminus rejected the request (HTTP {reply.status}): {self._text(reply)}",
                reply.status,
                self._text(reply),
            )
        body = self._json(reply)
        return body.get("data", body) if isinstance(body, dict) else body

    # ---- the calls this app needs

    def models(self) -> list[dict]:
        return self.call("GET", "/api/models")

    def devices(self) -> list[dict]:
        return self.call("GET", "/api/devices")

    def playlists(self) -> list[dict]:
        return self.call("GET", "/api/playlists")

    def screens(self) -> list[dict]:
        return self.call("GET", "/api/screens")

    def get_screen(self, screen_id: int) -> dict:
        return self.call("GET", f"/api/screens/{screen_id}")

    def create_screen(self, model_id: int, name: str, label: str, uri: str) -> dict:
        return self.call(
            "POST",
            "/api/screens",
            {
                "screen": {
                    "model_id": str(model_id),
                    "label": label,
                    "name": name,
                    "file_name": f"{name}.png",
                    "uri": uri,
                    "preprocessed": True,
                }
            },
        )

    def patch_screen(self, screen_id: int, uri: str) -> dict:
        return self.call(
            "PATCH", f"/api/screens/{screen_id}", {"screen": {"preprocessed": True, "uri": uri}}
        )

    def find_model(self, name: str) -> dict:
        for m in self.models():
            if m.get("name") == name:
                return m
        raise TerminusError(
            "config", f"Terminus has no model called '{name}'. Check TERMINUS_MODEL_NAME."
        )
