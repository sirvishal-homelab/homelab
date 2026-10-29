"""Battery / Wi-Fi readings from Terminus's device record (migration 0008)."""

from __future__ import annotations

import sqlite3
from datetime import UTC, datetime, timedelta

KEEP_DAYS = 90
MIN_GAP_MINUTES = 5
_FMT = "%Y-%m-%dT%H:%M:%SZ"


def _utc_text(value: object) -> str | None:
    """Terminus times come as ISO 8601 with an offset or Z; store them as UTC text."""
    if not isinstance(value, str) or not value:
        return None
    try:
        t = datetime.fromisoformat(value)
    except ValueError:
        return None
    if t.tzinfo is None:
        t = t.replace(tzinfo=UTC)
    return t.astimezone(UTC).strftime(_FMT)


def pick_device(devices: list[dict], model_id: int | None) -> dict | None:
    mine = [d for d in devices if model_id is not None and d.get("model_id") == model_id]
    return (mine or devices or [None])[0]


def record(
    conn: sqlite3.Connection,
    devices: list[dict],
    model_id: int | None,
    now: datetime | None = None,
) -> bool:
    """Store a reading if the device reported a voltage and has checked in since the last one."""
    device = pick_device(devices, model_id)
    volts = device.get("battery_voltage") if device else None
    if isinstance(volts, bool) or not isinstance(volts, int | float) or not 2.0 <= volts <= 5.5:
        return False
    now = now or datetime.now(UTC)
    seen = _utc_text(device.get("synced_at") or device.get("updated_at"))
    last = conn.execute(
        "SELECT at, seen_at FROM device_samples ORDER BY at DESC LIMIT 1"
    ).fetchone()
    if last:
        if seen and last["seen_at"] == seen:
            return False  # the device has not checked in again: nothing new to learn
        took = datetime.strptime(last["at"], _FMT).replace(tzinfo=UTC)
        if now - took < timedelta(minutes=MIN_GAP_MINUTES):
            return False
    wifi = device.get("wifi_signal")
    wifi = wifi if isinstance(wifi, int) and not isinstance(wifi, bool) else None
    conn.execute(
        "INSERT OR REPLACE INTO device_samples (at, voltage, wifi, seen_at) VALUES (?, ?, ?, ?)",
        (now.strftime(_FMT), round(float(volts), 3), wifi, seen),
    )
    cutoff = (now - timedelta(days=KEEP_DAYS)).strftime(_FMT)
    conn.execute("DELETE FROM device_samples WHERE at < ?", (cutoff,))
    conn.commit()
    return True


def since(conn: sqlite3.Connection, start: datetime) -> list[dict]:
    rows = conn.execute(
        "SELECT at, voltage, wifi, seen_at FROM device_samples WHERE at >= ? ORDER BY at",
        (start.astimezone(UTC).strftime(_FMT),),
    ).fetchall()
    return [dict(r) for r in rows]
