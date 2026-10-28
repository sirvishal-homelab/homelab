"""The connection check: is Terminus there, does the login work, is the device found, can the
image be fetched, is the screen in the device's playlist? Plain-language results (docs/05)."""

from __future__ import annotations

import ipaddress
import struct
import urllib.error
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from urllib.parse import urlparse

from trmnl_calendar.config import Settings
from trmnl_calendar.domain.models import Display
from trmnl_calendar.publish.terminus import TerminusClient, TerminusError

PNG = b"\x89PNG\r\n\x1a\n"


@dataclass(frozen=True)
class Step:
    key: str
    title: str
    status: str  # ok | warn | fail | info | skip
    detail: str = ""
    fix: str = ""


def _fetch(url: str) -> tuple[int, bytes]:
    try:
        with urllib.request.urlopen(url, timeout=10) as r:  # noqa: S310
            return r.status, r.read()
    except urllib.error.HTTPError as exc:
        return exc.code, b""
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise TerminusError("unreachable", str(getattr(exc, "reason", exc))) from exc


def _is_ip(host: str) -> bool:
    try:
        ipaddress.ip_address(host)
        return True
    except ValueError:
        return False


def missing_settings(s: Settings) -> list[str]:
    pairs = (
        ("TERMINUS_URL", s.terminus_url),
        ("TERMINUS_LOGIN", s.terminus_login),
        ("TERMINUS_PASSWORD", s.terminus_password),
        ("PUBLIC_BASE_URL", s.public_base_url),
    )
    return [k for k, v in pairs if not v]


def run_check(
    settings: Settings,
    display: Display,
    screen_id: int | None,
    fetch: Callable[[str], tuple[int, bytes]] = _fetch,
    client: TerminusClient | None = None,
) -> list[Step]:
    steps: list[Step] = []
    missing = missing_settings(settings)
    if missing:
        return [
            Step(
                "settings",
                "Settings",
                "fail",
                "Missing: " + ", ".join(missing),
                "Add them to the .env file on the server, then run "
                "`docker compose up -d` again to restart the app.",
            )
        ]
    host = urlparse(settings.public_base_url).hostname or ""
    if urlparse(settings.public_base_url).scheme not in ("http", "https") or not host:
        steps.append(
            Step(
                "settings",
                "Settings",
                "fail",
                "PUBLIC_BASE_URL is not a web address.",
                "Use a full address like http://192.168.1.20:8080 "
                "(the app's address as Terminus sees it).",
            )
        )
        return steps
    if _is_ip(host):
        steps.append(Step("settings", "Settings", "ok", "All four settings are present."))
    else:
        steps.append(
            Step(
                "settings",
                "Settings",
                "warn",
                f"PUBLIC_BASE_URL uses a name ({host}), not an IP address.",
                "Use the server's IP address, so a DNS outage can't stop pushes.",
            )
        )

    c = client or TerminusClient(
        settings.terminus_url,
        settings.terminus_login,
        settings.terminus_password,
        verify_tls=settings.terminus_verify_tls,
        timeout=10,
    )
    try:
        c.login()
        steps.append(
            Step("login", "Terminus and login", "ok", f"Signed in to {settings.terminus_url}.")
        )
    except TerminusError as exc:
        fix = (
            "Check TERMINUS_LOGIN and TERMINUS_PASSWORD in .env."
            if exc.kind == "auth"
            else "Check TERMINUS_URL, that Terminus is running, and that this server can reach it."
        )
        steps.append(Step("login", "Terminus and login", "fail", exc.message, fix))
        return steps + [Step(k, t, "skip", "Waiting for the login to work.") for k, t in _LATER]

    model = None
    try:
        model = c.find_model(settings.terminus_model_name)
        mine = (display.width, display.height, display.bit_depth)
        theirs = (model.get("width"), model.get("height"), model.get("bit_depth"))
        label = (
            f"{model.get('label', model['name'])} ({model['name']}): "
            f"{theirs[0]} x {theirs[1]}, {theirs[2]}-bit"
        )
        if mine == theirs:
            steps.append(Step("model", "Screen model", "ok", label))
        else:
            steps.append(
                Step(
                    "model",
                    "Screen model",
                    "fail",
                    f"{label}, but this app draws {mine[0]} x {mine[1]}, {mine[2]}-bit.",
                    "Set TERMINUS_MODEL_NAME to the model your device uses.",
                )
            )
    except TerminusError as exc:
        steps.append(
            Step("model", "Screen model", "fail", exc.message, "Fix TERMINUS_MODEL_NAME in .env.")
        )

    device = None
    try:
        devices = c.devices()
        if not devices:
            steps.append(
                Step(
                    "device",
                    "Device",
                    "fail",
                    "Terminus has no devices.",
                    "Set the TRMNL up with Terminus first.",
                )
            )
        else:
            mine = [d for d in devices if model and d.get("model_id") == model.get("id")]
            device = (mine or [None])[0]
            if device:
                bits = [device.get("label", "Device")]
                if device.get("firmware_version"):
                    bits.append(f"firmware {device['firmware_version']}")
                if device.get("battery_voltage") is not None:
                    bits.append(f"battery {device['battery_voltage']} V")
                if device.get("wifi_signal") is not None:
                    bits.append(f"Wi-Fi {device['wifi_signal']} dBm")
                if device.get("synced_at"):
                    bits.append(f"last check-in {device['synced_at']}")
                steps.append(Step("device", "Device", "ok", " · ".join(bits)))
            else:
                steps.append(
                    Step(
                        "device",
                        "Device",
                        "warn",
                        f"Terminus lists {len(devices)} device(s), none uses this model.",
                        "Check TERMINUS_MODEL_NAME matches the device's model in Terminus.",
                    )
                )
    except TerminusError as exc:
        steps.append(Step("device", "Device", "fail", exc.message))

    url = f"{settings.public_base_url}/render/current.png"
    try:
        status, body = fetch(url)
        if status == 200 and body.startswith(PNG):
            w, h = struct.unpack(">II", body[16:24])
            steps.append(Step("image", "Picture address", "ok", f"{url} returned a {w} x {h} PNG."))
        else:
            steps.append(
                Step(
                    "image",
                    "Picture address",
                    "warn",
                    f"{url} returned HTTP {status} and not a PNG.",
                    "Check PUBLIC_BASE_URL and HOST_PORT. Terminus must reach this address.",
                )
            )
    except TerminusError as exc:
        steps.append(
            Step(
                "image",
                "Picture address",
                "warn",
                f"This app could not open {url} ({exc.message}).",
                "Some networks block a container from reaching its own host address. "
                "**Push now** is the real test: Terminus does the fetch.",
            )
        )

    steps.append(_screen_step(c, screen_id, device, settings.terminus_screen_name))
    return steps


_LATER = (
    ("model", "Screen model"),
    ("device", "Device"),
    ("image", "Picture address"),
    ("screen", "Screen and playlist"),
)


def _screen_step(c: TerminusClient, screen_id: int | None, device: dict | None, name: str) -> Step:
    title = "Screen and playlist"
    if screen_id is None:
        return Step(
            "screen",
            title,
            "info",
            f"The screen '{name}' is created on the first push.",
            "After that, put it in the device's playlist in Terminus, as the only item.",
        )
    try:
        c.get_screen(screen_id)
    except TerminusError as exc:
        if exc.kind == "not_found":
            return Step(
                "screen",
                title,
                "warn",
                f"Screen {screen_id} is gone from Terminus.",
                "The next push creates it again; add the new one to the device's playlist.",
            )
        return Step("screen", title, "warn", exc.message)
    if not device or device.get("playlist_id") is None:
        return Step("screen", title, "ok", f"Screen {screen_id} exists in Terminus.")
    try:
        lists = c.playlists()
    except TerminusError as exc:
        return Step("screen", title, "warn", exc.message)
    pl = next((p for p in lists if p.get("id") == device["playlist_id"]), None)
    if pl is None:
        return Step("screen", title, "ok", f"Screen {screen_id} exists in Terminus.")
    ids = [i.get("screen_id") for i in pl.get("items", [])]
    label = pl.get("label") or pl.get("name")
    if screen_id not in ids:
        return Step(
            "screen",
            title,
            "warn",
            f"Screen {screen_id} is not in the device's playlist '{label}'.",
            "In Terminus, add the screen to that playlist, so the device shows it.",
        )
    if len(ids) > 1:
        return Step(
            "screen",
            title,
            "warn",
            f"The playlist '{label}' also has {len(ids) - 1} other screen(s); "
            "the device will alternate.",
            "Remove the other screens from that playlist if the calendar should always show.",
        )
    return Step(
        "screen", title, "ok", f"Screen {screen_id} is the only item in the playlist '{label}'."
    )
