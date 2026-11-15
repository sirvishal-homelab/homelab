from datetime import UTC, datetime, timedelta

from esp_screen.domain.alerts import alerts_for, card_for, track_since
from esp_screen.domain.model import Status, Tile

T0 = datetime(2026, 11, 15, 12, 0, tzinfo=UTC)


def status(down_systems=(), down_containers=()):
    return Status(
        rows=(),
        containers=Tile(0, 0),
        others=Tile(0, 0),
        down_systems=tuple(down_systems),
        down_containers=tuple(down_containers),
    )


def test_all_clear_has_no_card():
    alerts = alerts_for(status())
    assert alerts == []
    assert card_for(alerts, {}) is None


def test_system_down_card():
    alerts = alerts_for(status(["node-b"]))
    card = card_for(alerts, track_since(alerts, {}, T0))
    assert (card.title, card.also, card.since) == ("node-b down", "", T0)


def test_one_container_named_with_host():
    alerts = alerts_for(status(down_containers=["docker-host/app-a"]))
    card = card_for(alerts, track_since(alerts, {}, T0))
    assert card.title == "app-a stopped"
    assert card.detail == "container on docker-host"


def test_many_containers_grouped():
    alerts = alerts_for(status(down_containers=["h1/a", "h1/b", "h2/c"]))
    assert len(alerts) == 1
    assert alerts[0].title == "3 containers gone"
    assert alerts[0].detail == "on h1, h2"


def test_system_beats_container_and_summary_line():
    alerts = alerts_for(status(["node-b", "vm-1"], ["h/a"]))
    since = track_since(alerts, {}, T0)
    card = card_for(alerts, since)
    assert card.title == "node-b down"
    assert card.also == "also: 2 more problems"
    two = alerts_for(status(["node-b"], ["h/a"]))
    assert card_for(two, track_since(two, {}, T0)).also == "also: a stopped"


def test_since_is_kept_while_active_and_dropped_when_cleared():
    a1 = alerts_for(status(["node-b"]))
    s1 = track_since(a1, {}, T0)
    s2 = track_since(a1, s1, T0 + timedelta(minutes=5))
    assert s2["system:node-b"] == T0
    assert track_since([], s2, T0 + timedelta(minutes=6)) == {}


def test_oldest_of_same_severity_first():
    alerts = alerts_for(status(["a", "b"]))
    since = {"system:a": T0 + timedelta(minutes=3), "system:b": T0}
    assert card_for(alerts, since).title == "b down"
