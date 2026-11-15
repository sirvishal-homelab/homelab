from esp_screen.web.forms import parse

GOOD = {
    "warn_pct": ["75"],
    "crit_pct": ["90"],
    "poll_s": ["60"],
    "alert_flip_s": ["10"],
    "unreachable_grace_min": ["5"],
    "container_missing_polls": ["2"],
    "device_url": ["http://screen.test/"],
}
SYSTEMS = {"a", "b", "c", "d", "e"}


def form(**extra):
    f = {k: list(v) for k, v in GOOD.items()}
    for k, v in extra.items():
        f[k] = v if isinstance(v, list) else [v]
    return f


def test_defaults_parse():
    p = parse(form(), SYSTEMS)
    assert p.ok, p.errors
    assert p.numbers["poll_s"] == 60
    assert p.device_url == "http://screen.test"
    assert all(v == (None, None, False) for v in p.systems.values())


def test_focus_positions_names_and_ignore():
    p = parse(form(pos_a="2", pos_c="1", name_c=" NAS ", ignore_b="on"), SYSTEMS)
    assert p.ok, p.errors
    assert p.systems["a"] == (2, None, False)
    assert p.systems["c"] == (1, "NAS", False)
    assert p.systems["b"] == (None, None, True)


def test_gaps_in_positions_are_closed_in_order():
    p = parse(form(pos_a="4", pos_b="2"), SYSTEMS)
    assert p.ok
    assert (p.systems["b"][0], p.systems["a"][0]) == (1, 2)


def test_duplicate_position_rejected():
    p = parse(form(pos_a="1", pos_b="1"), SYSTEMS)
    assert "Two systems are set to position 1" in p.errors


def test_position_out_of_range_rejected():
    assert not parse(form(pos_a="5"), SYSTEMS).ok
    assert not parse(form(pos_a="x"), SYSTEMS).ok


def test_unknown_systems_in_form_are_ignored():
    p = parse(form(pos_zzz="1"), SYSTEMS)
    assert p.ok
    assert "zzz" not in p.systems


def test_number_errors():
    p = parse(form(poll_s="5", alert_flip_s="ten", warn_pct="95"), SYSTEMS)
    joined = " | ".join(p.errors)
    assert "between 15 and 600" in joined
    assert "whole number" in joined
    assert "amber must be below bar red" in joined


def test_screen_address():
    assert parse(form(device_url="192.0.2.5"), SYSTEMS).device_url == "http://192.0.2.5"
    assert not parse(form(device_url="https://screen.test"), SYSTEMS).ok
    assert not parse(form(device_url="http://screen.test/page"), SYSTEMS).ok
    assert parse(form(device_url=""), SYSTEMS).ok


def test_display_name_trimmed_to_twenty():
    p = parse(form(pos_a="1", name_a="x" * 40), SYSTEMS)
    assert p.systems["a"][1] == "x" * 20


def test_allowed_containers():
    p = parse(form(allow=["sys1:watchtower", "bad value", "sys2:app.web-1"]), SYSTEMS)
    assert p.allowed == {("sys1", "watchtower"), ("sys2", "app.web-1")}
