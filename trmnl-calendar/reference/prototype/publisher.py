"""Pushes the calendar image to a Terminus (TRMNL BYOS) server.

How it works
------------
Terminus fetches an image address when you create or update a screen through its API, and
does not re-fetch it on its own. So this module:

1. logs in to Terminus (POST /login) to get a short-lived token,
2. creates a screen once (POST /api/screens) pointing at this app's /calendar.png and
   remembers the screen id in the database,
3. updates that same screen (PATCH /api/screens/<id>) whenever something changes and on a
   timer, which makes Terminus fetch a freshly drawn image.

Everything is configured with environment variables (credentials stay out of the database).
"""
import json
import os
import ssl
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime

import store


def _env(name, default=""):
    return os.environ.get(name, default).strip()


def _flag(name, default):
    return _env(name, "true" if default else "false").lower() in ("1", "true", "yes", "on")


class PushError(Exception):
    pass


class Publisher:
    def __init__(self):
        self.base_url = _env("TERMINUS_URL").rstrip("/")
        self.login = _env("TERMINUS_LOGIN")
        self.password = os.environ.get("TERMINUS_PASSWORD", "")
        self.image_url = _env("TERMINUS_IMAGE_URL")
        self.model_id = _env("TERMINUS_MODEL_ID")
        self.playlist_id = _env("TERMINUS_PLAYLIST_ID")
        self.screen_name = _env("TERMINUS_SCREEN_NAME", "calendar")
        self.preprocessed = _flag("TERMINUS_PREPROCESSED", False)
        self.verify_tls = _flag("TERMINUS_VERIFY_TLS", True)
        try:
            self.interval = max(60, int(_env("TERMINUS_PUSH_INTERVAL", "900")))
        except ValueError:
            self.interval = 900
        self.debounce = 3.0

        self._scheme = None          # "raw" or "bearer", learned on first successful call
        self._lock = threading.Lock()
        self._wake = threading.Event()
        self._thread = None
        self._reason = "change"
        self._status = {"state": "idle", "message": "No push yet.", "at": None, "screen_id": None}

    # ------------------------------------------------------------ configuration

    def missing(self):
        needed = {
            "TERMINUS_URL": self.base_url,
            "TERMINUS_LOGIN": self.login,
            "TERMINUS_PASSWORD": self.password,
            "TERMINUS_IMAGE_URL": self.image_url,
        }
        return [k for k, v in needed.items() if not v]

    def configured(self):
        return not self.missing()

    def status(self):
        info = dict(self._status)
        info["configured"] = self.configured()
        info["missing"] = self.missing()
        info["terminus_url"] = self.base_url or None
        info["image_url"] = self.image_url or None
        info["interval_seconds"] = self.interval
        info["screen_id"] = store.get_internal("terminus_screen_id")
        return info

    # --------------------------------------------------------------------- HTTP

    def _request(self, method, path, body=None, token=None):
        url = self.base_url + path
        data = json.dumps(body).encode() if body is not None else None
        headers = {"Content-Type": "application/json", "Accept": "application/json"}
        if token:
            headers["Authorization"] = token if self._scheme != "bearer" else f"Bearer {token}"
        req = urllib.request.Request(url, data=data, method=method, headers=headers)
        ctx = None
        if url.startswith("https") and not self.verify_tls:
            ctx = ssl._create_unverified_context()
        try:
            with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
                raw = resp.read().decode() or "{}"
                return resp.status, _json(raw)
        except urllib.error.HTTPError as err:
            raw = err.read().decode() if err.fp else ""
            return err.code, _json(raw)
        except (urllib.error.URLError, TimeoutError, OSError) as err:
            raise PushError(f"Could not reach Terminus at {self.base_url}: {getattr(err, 'reason', err)}")

    def _login(self):
        status, body = self._request("POST", "/login", {"login": self.login, "password": self.password})
        token = (body or {}).get("access_token")
        if status >= 300 or not token:
            detail = (body or {}).get("error") or (body or {}).get("message") or f"HTTP {status}"
            raise PushError(f"Terminus login failed ({detail}). Check TERMINUS_LOGIN and TERMINUS_PASSWORD.")
        return token

    def _authed(self, method, path, body, token):
        """Terminus's docs show the bare token in the Authorization header; fall back to Bearer."""
        status, resp = self._request(method, path, body, token)
        if status == 401 and self._scheme is None:
            self._scheme = "bearer"
            status, resp = self._request(method, path, body, token)
            if status == 401:
                self._scheme = None
        elif status != 401 and self._scheme is None:
            self._scheme = "raw"
        return status, resp

    # --------------------------------------------------------------------- push

    def _screen_payload(self, creating):
        screen = {"uri": self.image_url}
        if self.preprocessed:
            screen["preprocessed"] = True
        if creating:
            screen.update({
                "label": "Calendar",
                "name": self.screen_name,
                "file_name": f"{self.screen_name}.png",
            })
            if self.model_id:
                screen["model_id"] = self.model_id
        body = {"screen": screen}
        return body

    def push(self, reason="manual"):
        """Send the current image to Terminus. Returns the status dict; never raises."""
        if not self.configured():
            self._set("error", "Terminus is not configured. Missing: " + ", ".join(self.missing()), reason)
            return self.status()
        with self._lock:
            try:
                token = self._login()
                screen_id = store.get_internal("terminus_screen_id")
                if screen_id:
                    status, resp = self._authed("PATCH", f"/api/screens/{screen_id}", self._screen_payload(False), token)
                    if status == 404:
                        screen_id = None          # screen was deleted in Terminus; recreate it
                if not screen_id:
                    if not self.model_id:
                        raise PushError("TERMINUS_MODEL_ID is required to create the screen the first time.")
                    body = self._screen_payload(True)
                    if self.playlist_id:
                        body["screen"]["playlist_id"] = self.playlist_id
                    status, resp = self._authed("POST", "/api/screens", body, token)
                if status == 401:
                    raise PushError("Terminus rejected the login token (HTTP 401).")
                if status >= 300:
                    detail = _problem(resp) or f"HTTP {status}"
                    raise PushError(f"Terminus refused the screen update: {detail}")
                new_id = ((resp or {}).get("data") or {}).get("id")
                if new_id:
                    store.set_internal("terminus_screen_id", new_id)
                self._set("ok", "Sent the latest calendar image to Terminus.", reason)
            except PushError as err:
                self._set("error", str(err), reason)
            except Exception as err:  # keep the worker alive whatever happens
                self._set("error", f"Unexpected error: {err}", reason)
        return self.status()

    def _set(self, state, message, reason):
        self._status = {"state": state, "message": message, "reason": reason,
                        "at": datetime.now().isoformat(timespec="seconds")}

    # ------------------------------------------------------------------- worker

    def request_push(self, reason="change"):
        """Ask the background worker to push soon (debounced). Safe to call from request handlers."""
        if self.configured():
            self._reason = reason
            self._wake.set()

    def start(self):
        if self._thread or not self.configured():
            return
        self._thread = threading.Thread(target=self._loop, name="terminus-publisher", daemon=True)
        self._thread.start()

    def _loop(self):
        time.sleep(5)                       # let the web server finish starting
        self.push("startup")
        while True:
            if self._wake.wait(timeout=self.interval):
                time.sleep(self.debounce)   # collect a burst of edits into one push
                self._wake.clear()
                self.push(self._reason)
            else:
                self.push("schedule")


def _json(raw):
    try:
        return json.loads(raw) if raw else {}
    except ValueError:
        return {}


def _problem(resp):
    """Pull a readable message out of an RFC 7807-style error body."""
    if not isinstance(resp, dict):
        return ""
    return str(resp.get("detail") or resp.get("title") or resp.get("error") or resp.get("message") or "")


publisher = Publisher()
