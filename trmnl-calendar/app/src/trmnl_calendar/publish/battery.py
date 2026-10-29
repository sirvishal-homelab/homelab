"""Turn battery voltage readings into something readable. Pure: no I/O, no clock.

Terminus reports a voltage, not a percentage. The percentage here is an estimate from a typical
single-cell lithium-polymer discharge curve, so treat it as approximate. The trend (is it going
down, and how fast) is more trustworthy than any single number.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta

# (volts, percent), highest first. A common LiPo curve; the middle of the range is nearly flat.
CURVE = (
    (4.20, 100),
    (4.15, 95),
    (4.11, 90),
    (4.08, 85),
    (4.02, 80),
    (3.98, 75),
    (3.95, 70),
    (3.91, 65),
    (3.87, 60),
    (3.85, 55),
    (3.84, 50),
    (3.82, 45),
    (3.80, 40),
    (3.79, 35),
    (3.77, 30),
    (3.75, 25),
    (3.73, 20),
    (3.71, 15),
    (3.69, 10),
    (3.61, 5),
    (3.27, 0),
)
RISE_V = 0.03  # a rise this big between readings means it is charging
FULL_V = 4.15
_FMT = "%Y-%m-%dT%H:%M:%SZ"


def estimate_percent(volts: float) -> int:
    if volts >= CURVE[0][0]:
        return 100
    if volts <= CURVE[-1][0]:
        return 0
    for (v_hi, p_hi), (v_lo, p_lo) in zip(CURVE, CURVE[1:], strict=False):
        if v_lo <= volts <= v_hi:
            return round(p_lo + (volts - v_lo) / (v_hi - v_lo) * (p_hi - p_lo))
    return 0


def wifi_label(dbm: int | None) -> str:
    if dbm is None:
        return ""
    if dbm >= -60:
        return "strong"
    if dbm >= -70:
        return "good"
    if dbm >= -80:
        return "weak"
    return "very weak"


def _t(text: str) -> datetime:
    return datetime.strptime(text, _FMT).replace(tzinfo=UTC)


def _state(samples: list[dict], now: datetime) -> str:
    """charging | full | draining | unknown, from the last six hours of readings."""
    recent = [s for s in samples if _t(s["at"]) >= now - timedelta(hours=6)]
    if len(recent) < 2 or _t(recent[-1]["at"]) - _t(recent[0]["at"]) < timedelta(hours=1):
        return "unknown"
    volts = [s["voltage"] for s in recent]
    if volts[-1] - min(volts) >= RISE_V:
        return "charging"
    if min(volts) >= FULL_V and max(volts) - min(volts) < 0.05:
        return "full"
    if volts[0] - volts[-1] >= 0.01:
        return "draining"
    return "unknown"


def _drain(samples: list[dict]) -> float | None:
    """Percent lost per day over the latest stretch with no charging in it, or None."""
    start = 0
    for i in range(1, len(samples)):
        if samples[i]["voltage"] - samples[i - 1]["voltage"] >= 0.05:
            start = i  # charged here: only look at what came after
    window = samples[start:]
    if len(window) < 2:
        return None
    span = (_t(window[-1]["at"]) - _t(window[0]["at"])).total_seconds() / 86400
    lost = estimate_percent(window[0]["voltage"]) - estimate_percent(window[-1]["voltage"])
    if span < 0.5 or lost < 2:
        return None
    return lost / span


def sparkline(samples: list[dict], now: datetime, days: int = 7, w: int = 300, h: int = 60) -> str:
    """SVG polyline points for the last `days` of voltage, on a fixed 3.2 to 4.3 V scale."""
    lo, hi = now - timedelta(days=days), now
    pts = []
    for s in samples:
        t = _t(s["at"])
        if t < lo:
            continue
        x = (t - lo).total_seconds() / (hi - lo).total_seconds() * w
        y = h - (min(max(s["voltage"], 3.2), 4.3) - 3.2) / 1.1 * h
        pts.append(f"{x:.1f},{y:.1f}")
    return " ".join(pts)


def describe(samples: list[dict], now: datetime, refresh_seconds: int | None = None) -> dict | None:
    """Summary for Settings. `samples` are oldest first. None when there are no readings yet."""
    if not samples:
        return None
    last = samples[-1]
    volts = last["voltage"]
    percent = estimate_percent(volts)
    state = _state(samples, now)
    per_day = _drain(samples) if state != "charging" else None
    if state == "unknown" and per_day:
        state = "draining"  # too slow to see in six hours, but clear over days
    days_left = round(percent / per_day, 1) if per_day and per_day > 0 else None
    seen = last.get("seen_at") or last["at"]
    silent_minutes = int((now - _t(seen)).total_seconds() // 60)
    limit = max(3 * (refresh_seconds or 900) // 60, 30)
    return {
        "voltage": round(volts, 2),
        "percent": percent,
        "wifi": last.get("wifi"),
        "wifi_label": wifi_label(last.get("wifi")),
        "state": state,
        "per_day": round(per_day, 1) if per_day else None,
        "days_left": days_left,
        "seen_at": seen,
        "silent_minutes": silent_minutes,
        "stale": silent_minutes > limit,
        "readings": len(samples),
        "first_at": samples[0]["at"],
        "points": sparkline(samples, now),
    }
