import io
from datetime import UTC, date, datetime

import pytest
from PIL import Image

from trmnl_calendar.render.pipeline import RenderCache
from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.web import forms, render_service, views

# 01:00 UTC on 28 Oct is still the evening of 27 Oct in Halifax (UTC-3 until 1 Nov)
NOW = datetime(2026, 10, 28, 1, 0, tzinfo=UTC)


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def add(conn, **kw):
    d = dict(title="Dentist", section_id=1, start_date="2026-10-28")
    d.update(kw)
    return repo.create_event(conn, d)


def test_today_uses_the_calendar_timezone():
    assert render_service.today_for("America/Halifax", NOW) == date(2026, 10, 27)
    assert render_service.today_for("UTC", NOW) == date(2026, 10, 28)


def test_render_current_is_a_valid_png_and_is_cached(conn):
    add(conn, start_time="09:30")
    cache = RenderCache()
    r = render_service.render_current(conn, NOW, cache)
    assert cache.get(r.hash) == r.png
    im = Image.open(io.BytesIO(r.png))
    assert im.size == (800, 480)
    assert any(getattr(t, "text", "") == "9:30a Dentist" for t in r.scene.items)


def test_render_is_stable_until_content_changes(conn):
    ev = add(conn)
    a = render_service.render_current(conn, NOW)
    assert render_service.render_current(conn, NOW).hash == a.hash
    repo.update_event(conn, ev.id, {"title": "Dentist!"}, 1)
    assert render_service.render_current(conn, NOW).hash != a.hash


def test_skipped_and_deleted_events_do_not_render(conn):
    ev = add(conn, freq="daily")
    repo.add_skip(conn, ev.id, "2026-10-29")
    texts = [getattr(t, "text", "") for t in render_service.render_current(conn, NOW).scene.items]
    assert texts.count("Dentist") == 3  # 28, 30 and 31 Oct; the 29th is skipped
    repo.soft_delete(conn, ev.id, 1)
    texts = [getattr(t, "text", "") for t in render_service.render_current(conn, NOW).scene.items]
    assert "Dentist" not in texts


def test_home_lists_upcoming_with_labels_and_escapes_html(conn):
    add(conn, title="<script>x</script>", start_time="14:00", end_time="15:00")
    add(conn, title="Cabin", start_date="2026-10-30", end_date="2026-11-01")
    model = views.upcoming_groups(conn, NOW)
    labels = [label for _, label, _ in model["days"]]
    assert labels[:2] == ["Tomorrow", "Fri Oct 30"]
    html = views.render_page("home.html", flash="", flash_deleted=None, **model)
    assert "<script>x</script>" not in html and "&lt;script&gt;" in html
    assert "2p–3p" in html and "Day 1 of 3" in html and "/render/current.png" in html


def test_home_empty_state(conn):
    html = views.render_page(
        "home.html", flash="", flash_deleted=None, **views.upcoming_groups(conn, NOW)
    )
    assert "Nothing in the next 28 days" in html


def test_form_pages_render_prefilled_and_with_errors(conn):
    ev = add(conn, freq="weekly", interval=2, location="Arena")
    values = forms.event_to_values(ev)
    html = views.render_page("event_form.html", **views.form_context(conn, values, event=ev))
    assert 'value="Dentist"' in html and "Repeats every 2 weeks" in html
    assert 'name="rev" value="1"' in html and "Days off" in html and "Delete this event" in html
    bad = forms.event_to_values(None, "2026-10-28", 1)
    html = views.render_page(
        "event_form.html", **views.form_context(conn, bad, {"title": "Enter a short description"})
    )
    assert "Enter a short description" in html and 'role="alert"' in html
