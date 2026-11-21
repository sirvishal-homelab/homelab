"""Builds the settings page context and HTML (no web framework here, so it is tested anywhere)."""

from __future__ import annotations

import sqlite3
import time
from datetime import UTC, datetime
from importlib import resources
from zoneinfo import ZoneInfo

from jinja2 import Environment, FunctionLoader, select_autoescape

from esp_screen import __version__
from esp_screen.storage import repo
from esp_screen.web.forms import NUMBERS

DEFAULTS = {
    "warn_pct": 75,
    "crit_pct": 90,
    "poll_s": 60,
    "alert_flip_s": 10,
    "unreachable_grace_min": 5,
    "container_missing_polls": 2,
}


def _load(name: str) -> str:
    return (resources.files("esp_screen.web") / "templates" / name).read_text(encoding="utf-8")


_env = Environment(loader=FunctionLoader(_load), autoescape=select_autoescape(["html"]))


def local_time(iso: str | None, tz: str, now: datetime | None = None) -> str:
    """'2026-11-15T16:21:34Z' -> '12:21' (24-hour, local); the date is added if not today."""
    if not iso:
        return "-"
    try:
        when = datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone(ZoneInfo(tz))
    except ValueError:
        return iso
    today = (now or datetime.now(UTC)).astimezone(ZoneInfo(tz)).date()
    return when.strftime("%H:%M") if when.date() == today else when.strftime("%Y-%m-%d %H:%M")


def render(
    conn: sqlite3.Connection,
    runner,
    *,
    flags: dict[str, bool] | None = None,
    errors: list[str] | None = None,
    warnings: list[str] | None = None,
    tz: str = "America/Halifax",
) -> str:
    flags = flags or {}
    beszel = runner.health if runner else None
    alert = runner.plan.card if runner else None
    has_alert = bool(runner and runner.last_alert)
    states = dict(runner.last_states) if runner else {}
    push = repo.get_push_state(conn)
    return _env.get_template("settings.html").render(
        last_push=local_time(push.last_push_at, tz),
        last_read=local_time(beszel.last_ok_at if beszel else None, tz),
        version=__version__,
        now_ts=int(time.time()),
        saved=flags.get("saved"),
        pushed=flags.get("pushed"),
        forgot=flags.get("forgot"),
        errors=errors or [],
        warnings=warnings or [],
        push=push,
        beszel=beszel,
        alert=alert,
        has_alert=has_alert,
        systems=repo.list_systems(conn),
        states=states,
        containers=repo.list_containers(conn),
        numbers=NUMBERS,
        values={k: repo.get_setting(conn, k, d) for k, d in DEFAULTS.items()},
        device_url=repo.get_setting(conn, "device_url", ""),
        show_updates=bool(repo.get_setting(conn, "show_updates", True)),
    )
