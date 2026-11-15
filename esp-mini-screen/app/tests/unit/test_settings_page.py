"""Renders the settings page HTML without a web server (the routes are tested in CI)."""

import contextlib
from types import SimpleNamespace

from esp_screen.domain.alerts import Card
from esp_screen.loop import BeszelHealth
from esp_screen.storage import db, repo
from esp_screen.web import page


def fake_runner(**kw):
    base = dict(
        health=BeszelHealth(last_ok_at="2026-11-15T16:00:00Z"),
        plan=SimpleNamespace(card=None),
        last_alert=None,
        last_states={"id-a": "up", "id-b": "down"},
    )
    base.update(kw)
    return SimpleNamespace(**base)


def html(settings, runner=None, **kw):
    with contextlib.closing(db.connect(settings.db_path)) as c:
        repo.sync_systems(c, [("id-a", "node-a"), ("id-b", "node-b")])
        repo.save_system_prefs(c, {"id-a": (1, "Main", False), "id-b": (None, None, True)})
        repo.save_container_memory(c, {("id-a", "web"): 0, ("id-a", "old"): 4})
        return page.render(c, runner, **kw)


def test_page_lists_systems_with_choices(migrated):
    out = html(migrated, fake_runner())
    assert 'name="pos_id-a"' in out and 'value="1" selected' in out
    assert 'value="Main"' in out
    assert 'name="ignore_id-b" value="1" checked' in out
    assert 'class="note state-down">down' in out


def test_page_lists_containers_and_forget(migrated):
    out = html(migrated, fake_runner())
    assert 'value="id-a:web"' in out
    assert "not running (4 reads)" in out
    assert 'action="/containers/forget"' in out
    assert 'value="id-a:old"' in out


def test_page_shows_alert_and_errors_escaped(migrated):
    card = Card("node-b down", "Beszel reports it down", None, "also: <b>x</b>")
    out = html(migrated, fake_runner(plan=SimpleNamespace(card=card)), errors=["<script>bad"])
    assert "node-b down" in out
    assert "&lt;script&gt;bad" in out and "<script>bad" not in out
    assert "&lt;b&gt;x&lt;/b&gt;" in out


def test_page_without_beszel(migrated):
    out = html(migrated, None, flags={"saved": True})
    assert "not configured (.env)" in out
    assert "Saved." in out


def test_tuning_fields_have_current_values(migrated):
    with contextlib.closing(db.connect(migrated.db_path)) as c:
        repo.set_setting(c, "poll_s", 120)
    out = html(migrated, fake_runner())
    assert 'name="poll_s" value="120"' in out


def test_local_time_24_hour():
    from datetime import UTC, datetime

    now = datetime(2026, 11, 15, 18, 0, tzinfo=UTC)
    assert page.local_time("2026-11-15T16:21:34Z", "America/Halifax", now) == "12:21"
    assert page.local_time("2026-11-14T16:21:34Z", "America/Halifax", now) == "2026-11-14 12:21"
    assert page.local_time(None, "America/Halifax", now) == "-"
