"""Settings form: parse and validate (pure; no web framework, so it is tested everywhere).

Field names: `pos_<id>` (""/1-4), `name_<id>`, `ignore_<id>` per system; `allow` (repeated,
"<system_id>:<container>") for containers allowed to stop; `show_updates` (checkbox); plus the
tuning fields below.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from urllib.parse import urlsplit

# key -> (label, minimum, maximum); all whole numbers
NUMBERS: dict[str, tuple[str, int, int]] = {
    "warn_pct": ("Bar amber at (%)", 1, 100),
    "crit_pct": ("Bar red at (%)", 1, 100),
    "poll_s": ("Read Beszel every (s)", 15, 600),
    "alert_flip_s": ("Alert flip every (s)", 3, 60),
    "unreachable_grace_min": ("Beszel unreachable turns red after (min)", 1, 120),
    "container_missing_polls": ("Container gone after missed reads", 1, 10),
}
MAX_FOCUS = 4
MAX_NAME = 20
_CONTAINER = re.compile(r"^([A-Za-z0-9_-]+):([A-Za-z0-9][A-Za-z0-9_.-]*)$")


@dataclass
class ParsedSettings:
    numbers: dict[str, int] = field(default_factory=dict)
    device_url: str = ""
    show_updates: bool = True
    systems: dict[str, tuple[int | None, str | None, bool]] = field(default_factory=dict)
    allowed: set[tuple[str, str]] = field(default_factory=set)
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors


def _one(form: dict[str, list[str]], key: str) -> str:
    values = form.get(key) or [""]
    return values[0].strip()


def parse(form: dict[str, list[str]], known_systems: set[str]) -> ParsedSettings:
    out = ParsedSettings()

    for key, (label, lo, hi) in NUMBERS.items():
        raw = _one(form, key)
        try:
            value = int(raw)
        except ValueError:
            out.errors.append(f"{label}: enter a whole number")
            continue
        if not lo <= value <= hi:
            out.errors.append(f"{label}: must be between {lo} and {hi}")
            continue
        out.numbers[key] = value
    if {"warn_pct", "crit_pct"} <= out.numbers.keys():
        if out.numbers["warn_pct"] >= out.numbers["crit_pct"]:
            out.errors.append("Bar amber must be below bar red")

    url = _one(form, "device_url").rstrip("/")
    if url:
        if "://" not in url:
            url = f"http://{url}"
        parts = urlsplit(url)
        if parts.scheme != "http" or not parts.hostname or parts.path not in ("", "/"):
            out.errors.append("Screen address: use http://<address> (the screen has no https)")
    out.device_url = url
    out.show_updates = _one(form, "show_updates") in ("1", "on", "true")

    positions: dict[int, str] = {}
    for sid in sorted(known_systems):
        raw = _one(form, f"pos_{sid}")
        order: int | None = None
        if raw:
            try:
                order = int(raw)
            except ValueError:
                order = 0
            if not 1 <= order <= MAX_FOCUS:
                out.errors.append(f"Position must be 1 to {MAX_FOCUS}")
                order = None
            elif order in positions:
                out.errors.append(f"Two systems are set to position {order}")
            else:
                positions[order] = sid
        name = _one(form, f"name_{sid}")[:MAX_NAME] or None
        ignore = _one(form, f"ignore_{sid}") in ("1", "on", "true")
        out.systems[sid] = (order, name, ignore)
    if positions and sorted(positions) != list(range(1, len(positions) + 1)):
        # Gaps are harmless but confusing; close them in the chosen order.
        for new, old in enumerate(sorted(positions), start=1):
            sid = positions[old]
            _, name, ignore = out.systems[sid]
            out.systems[sid] = (new, name, ignore)

    for raw in form.get("allow", []):
        m = _CONTAINER.match(raw.strip())
        if m:
            out.allowed.add((m.group(1), m.group(2)))
    return out
