"""Alert rules (docs/04 red conditions): pure, from Status plus first-seen times.

Red: a system down (paused never counts), an expected container gone. Beszel being unreachable
or its data changing is handled by the loop with its own frames, because there is no Status then.

The screen shows one alert card (the most severe, oldest first) with a one-line summary of the
rest, alternating with the overview. One card keeps uploads to a minimum: flipping between two
stored images needs no upload (docs/05 R-4).
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime

from esp_screen.domain.model import Status


@dataclass(frozen=True)
class Alert:
    key: str
    severity: int  # lower is more severe
    title: str
    detail: str


@dataclass(frozen=True)
class Card:
    title: str
    detail: str
    since: datetime
    also: str  # "" when there is nothing else


def alerts_for(status: Status) -> list[Alert]:
    out = [
        Alert(f"system:{n}", 0, f"{n} down", "Beszel reports it down") for n in status.down_systems
    ]
    gone = list(status.down_containers)
    if len(gone) == 1:
        system, name = gone[0].split("/", 1)
        out.append(Alert(f"container:{gone[0]}", 1, f"{name} stopped", f"container on {system}"))
    elif gone:
        hosts = sorted({g.split("/", 1)[0] for g in gone})
        out.append(
            Alert(
                "containers:" + ",".join(sorted(gone)),
                1,
                f"{len(gone)} containers gone",
                "on " + ", ".join(hosts),
            )
        )
    return out


def track_since(
    alerts: list[Alert], previous: Mapping[str, datetime], now: datetime
) -> dict[str, datetime]:
    """First-seen time per alert key; keys that cleared are dropped."""
    return {a.key: previous.get(a.key, now) for a in alerts}


def card_for(alerts: list[Alert], since: Mapping[str, datetime]) -> Card | None:
    if not alerts:
        return None
    ordered = sorted(alerts, key=lambda a: (a.severity, since[a.key], a.key))
    top, rest = ordered[0], ordered[1:]
    if not rest:
        also = ""
    elif len(rest) == 1:
        also = f"also: {rest[0].title}"
    else:
        also = f"also: {len(rest)} more problems"
    return Card(top.title, top.detail, since[top.key], also)
