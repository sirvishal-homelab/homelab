from datetime import UTC, datetime

import pytest

from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.web import views

NOW = datetime(2026, 3, 15, 15, 0, tzinfo=UTC)


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def add(conn, **kw):
    d = dict(title="Dentist", section_id=1, start_date="2026-03-10")
    d.update(kw)
    return repo.create_event(conn, d)


def cells(ctx):
    return {c["iso"]: c for row in ctx["rows"] for c in row}


def test_grid_is_whole_weeks_starting_on_week_start(conn):
    ctx = views.month_context(conn, NOW, 2026, 3)
    assert ctx["title"] == "March 2026"
    assert ctx["weekday_names"][0] == "Sun"
    assert all(len(r) == 7 for r in ctx["rows"])
    assert ctx["rows"][0][0]["iso"] == "2026-03-01"  # March 1, 2026 is a Sunday
    assert len(ctx["rows"]) == 5
    assert cells(ctx)["2026-03-15"]["today"] is True


def test_defaults_to_current_month_and_clamps_bad_input(conn):
    assert views.month_context(conn, NOW)["month"] == 3
    assert views.month_context(conn, NOW, 2026, 13)["month"] == 3
    assert views.month_context(conn, NOW, 2026, 12)["next"] == (2027, 1)
    assert views.month_context(conn, NOW, 2026, 1)["prev"] == (2025, 12)


def test_events_trips_and_overflow(conn):
    add(conn, title="A", start_date="2026-03-10", start_time="09:00")
    add(conn, title="Trip", start_date="2026-03-11", end_date="2026-03-13")
    for i in range(4):
        add(conn, title=f"E{i}", start_date="2026-03-20")
    c = cells(views.month_context(conn, NOW, 2026, 3))
    assert c["2026-03-10"]["events"][0]["label"].endswith("A")
    assert [len(c[d]["events"]) for d in ("2026-03-11", "2026-03-12", "2026-03-13")] == [1, 1, 1]
    assert c["2026-03-12"]["events"][0]["mark"] == "◂ ▸"
    assert c["2026-03-20"]["count"] == 4 and c["2026-03-20"]["more"] == 1


def test_selected_day_and_bad_day(conn):
    add(conn, title="A", start_date="2026-03-10")
    ctx = views.month_context(conn, NOW, 2026, 3, day="2026-03-10")
    assert ctx["selected_label"] == "Tue Mar 10"
    assert [i["title"] for i in ctx["selected_items"]] == ["A"]
    assert views.month_context(conn, NOW, 2026, 3, day="nope")["selected"] is None


def test_add_query_replaces_existing_param():
    out = views.add_query("/month?y=2026&m=3", saved="Hi there")
    assert out == "/month?y=2026&m=3&saved=Hi+there"
    assert views.add_query("/month?saved=old&m=3", saved="x") == "/month?m=3&saved=x"


def test_schedule_days_carry_the_note_and_action_fields(conn):
    rot = [["MO", "TU", "FR", "SA", "SU"], ["WE", "TH"]]
    ev = add(conn, title="Work", start_date="2026-03-02", rotation=rot, start_time="08:00")
    repo.set_day_note(conn, ev.id, "2026-03-06", "out")
    c = cells(views.month_context(conn, NOW, 2026, 3, day="2026-03-06"))
    fri = c["2026-03-06"]["events"][0]
    assert fri["label"].endswith("Work (out)") and fri["note"] == "out" and fri["scheduled"]
    assert fri["date"] == "2026-03-06"
    assert c["2026-03-05"]["events"] == []  # Thursday of week A is not a work day
    assert c["2026-03-11"]["events"][0]["note"] == ""  # Wednesday of week B, undecided
    home = views.upcoming_groups(conn, NOW)
    first = home["days"][0][2][0]
    assert first["scheduled"] and first["date"] and first["title"].startswith("Work")


def test_month_template_shows_day_buttons_only_for_schedules(conn):
    rot = [["MO"], ["TU"]]
    add(conn, title="Work", start_date="2026-03-16", rotation=rot)
    add(conn, title="Dentist", start_date="2026-03-16")
    ctx = views.month_context(conn, NOW, 2026, 3, day="2026-03-16")
    html = views.render_page(
        "month.html",
        **{
            **ctx,
            "me": None,
            "request_path": "/month",
            "nav": "home",
            "flash": "",
            "flash_deleted": None,
            "back": "/month?y=2026&m=3",
        },
    )
    assert html.count('action="/events/1/day"') == 2  # Out and In for the schedule
    assert 'action="/events/1/skip"' in html and "/events/2/day" not in html
