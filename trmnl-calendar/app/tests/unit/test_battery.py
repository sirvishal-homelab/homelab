from datetime import UTC, datetime, timedelta

from trmnl_calendar.publish import battery

NOW = datetime(2026, 10, 29, 18, 0, tzinfo=UTC)
FMT = "%Y-%m-%dT%H:%M:%SZ"


def series(points, **kw):
    """points: (hours_ago, volts). Returns readings oldest first."""
    out = []
    for hours_ago, v in sorted(points, reverse=True):
        t = (NOW - timedelta(hours=hours_ago)).strftime(FMT)
        out.append({"at": t, "voltage": v, "wifi": -55, "seen_at": t, **kw})
    return out


def test_percent_curve_is_monotonic_and_clamped():
    assert battery.estimate_percent(4.3) == 100 and battery.estimate_percent(4.2) == 100
    assert battery.estimate_percent(3.0) == 0 and battery.estimate_percent(3.27) == 0
    assert battery.estimate_percent(3.84) == 50
    last = 101
    for v in [x / 100 for x in range(430, 290, -1)]:
        p = battery.estimate_percent(v)
        assert p <= last
        last = p


def test_no_readings_gives_nothing():
    assert battery.describe([], NOW) is None


def test_draining_shows_rate_and_days_left():
    s = series([(72, 4.02), (48, 3.98), (24, 3.91), (0, 3.85)])
    s = series([(72, 4.02), (48, 3.98), (24, 3.91), (3, 3.86), (2, 3.855), (0, 3.85)])
    d = battery.describe(s, NOW, 600)
    assert d["percent"] == 55 and d["state"] == "draining"
    assert d["per_day"] and 5 < d["per_day"] < 12  # (80 -> 55) over 3 days
    assert d["days_left"] and 4 < d["days_left"] < 12
    assert not d["stale"] and d["readings"] == 6


def test_rising_voltage_means_charging_and_hides_the_drain_rate():
    s = series([(5, 3.80), (4, 3.82), (3, 3.88), (2, 3.95), (0, 4.05)])
    d = battery.describe(s, NOW)
    assert d["state"] == "charging" and d["per_day"] is None


def test_steady_near_full_means_probably_on_usb():
    s = series([(5, 4.18), (4, 4.19), (3, 4.18), (2, 4.19), (0, 4.19)])
    assert battery.describe(s, NOW)["state"] == "full"


def test_too_few_readings_says_unknown():
    s = series([(0.2, 3.9), (0, 3.9)])
    d = battery.describe(s, NOW)
    assert d["state"] == "unknown" and d["per_day"] is None


def test_drain_rate_ignores_everything_before_the_last_charge():
    s = series([(96, 3.70), (72, 3.65), (60, 4.15), (36, 4.02), (12, 3.91), (0, 3.85)])
    d = battery.describe(s, NOW)
    # from the charge at -60h: 4.15 (95%) -> 3.85 (55%) over 2.5 days, not the earlier dip
    assert d["per_day"] and 14 < d["per_day"] < 18


def test_device_that_stopped_checking_in_is_flagged():
    s = series([(8, 3.9), (7, 3.9)])
    d = battery.describe(s, NOW, 600)
    assert d["stale"] and d["silent_minutes"] >= 7 * 60 - 1


def test_wifi_words_and_sparkline():
    assert [battery.wifi_label(x) for x in (-50, -65, -75, -90, None)] == [
        "strong",
        "good",
        "weak",
        "very weak",
        "",
    ]
    pts = battery.sparkline(series([(24, 4.2), (0, 3.2)]), NOW).split()
    assert len(pts) == 2 and pts[0].endswith(",0.0") is False
    assert pts[1].split(",")[1] == "60.0"  # 3.2 V sits on the bottom edge
