"""The Shifts pages' models: every day of a rotating schedule is none, working, in, out or off."""

from datetime import UTC, date, datetime

import pytest

from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import people as people_repo
from trmnl_calendar.web import shifts_view

# Tuesday 3 Nov 2026 (the clocks went back on Sunday 1 Nov)
NOW = datetime(2026, 11, 3, 15, 0, tzinfo=UTC)
WEEK_A = ["WE", "TH"]
WEEK_B = ["MO", "TU", "FR", "SA", "SU"]


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def schedule(conn, **kw):
    d = dict(
        title="Work",
        section_id=2,
        start_date="2026-10-26",  # a Monday: Week A is 26 Oct, Week B is 2 Nov
        rotation=[WEEK_A, WEEK_B],
    )
    d.update(kw)
    return repo.create_event(conn, d)


def by_iso(week):
    return {d["iso"]: d for d in week["days"]}


def test_overview_strip_is_as_long_as_the_rotation_and_starts_on_this_monday(conn):
    schedule(conn)
    card = shifts_view.overview_context(conn, NOW)["cards"][0]
    assert [w["name"] for w in card["weeks"]] == ["B", "A"]  # this week (2 Nov) is week B
    assert card["this_week"] == "B"
    assert card["weeks"][0]["days"][0]["iso"] == "2026-11-02"
    three = schedule(conn, title="Three", rotation=[["MO"], ["TU"], ["WE"]])
    cards = {c["event_id"]: c for c in shifts_view.overview_context(conn, NOW)["cards"]}
    assert len(cards[three.id]["weeks"]) == 3


def test_states_working_in_out_and_day_off_are_distinct(conn):
    e = schedule(conn)
    repo.set_day_note(conn, e.id, "2026-11-03", "out")
    repo.set_day_note(conn, e.id, "2026-11-06", "in")
    repo.add_skip(conn, e.id, "2026-11-07")
    week = shifts_view.overview_context(conn, NOW)["cards"][0]["weeks"][0]
    days = by_iso(week)
    assert days["2026-11-02"]["state"] == "working"
    assert days["2026-11-03"]["state"] == "out" and days["2026-11-03"]["word"] == "out"
    assert days["2026-11-04"]["state"] == "none"  # a Wednesday in week B
    assert days["2026-11-06"]["state"] == "in"
    assert days["2026-11-07"]["state"] == "off" and days["2026-11-07"]["word"] == "off"
    assert days["2026-11-03"]["today"] is True


def test_a_day_off_wins_over_a_note_on_the_same_day(conn):
    e = schedule(conn)
    repo.set_day_note(conn, e.id, "2026-11-03", "in")
    repo.add_skip(conn, e.id, "2026-11-03")
    day = by_iso(shifts_view.overview_context(conn, NOW)["cards"][0]["weeks"][0])["2026-11-03"]
    assert day["state"] == "off"


def test_empty_week_and_no_schedules(conn):
    assert shifts_view.overview_context(conn, NOW)["cards"] == []
    ctx = shifts_view.month_context(conn, NOW)
    assert ctx["chosen"] is None and ctx["rows"] == []
    schedule(conn, rotation=[[], ["MO"]])
    week_a = shifts_view.overview_context(conn, NOW)["cards"][0]["weeks"][1]
    assert {d["state"] for d in week_a["days"]} == {"none"}


def test_month_grid_marks_the_pattern_and_offers_actions_for_the_chosen_day(conn):
    e = schedule(conn)
    repo.add_skip(conn, e.id, "2026-11-12")
    ctx = shifts_view.month_context(conn, NOW, e.id, 2026, 11, "2026-11-04")
    cells = {c["iso"]: c for row in ctx["rows"] for c in row}
    assert cells["2026-11-02"]["state"] == "working"
    assert cells["2026-11-04"]["state"] == "none"  # Wednesday of week B
    assert cells["2026-11-05"]["state"] == "none"
    assert cells["2026-11-11"]["state"] == "working"  # Wednesday of week A
    assert cells["2026-11-12"]["state"] == "off"
    assert ctx["title"] == "November 2026" and ctx["is_current"] is True
    panel = ctx["panel"]
    assert panel["state"] == "none" and panel["scheduled"] is False
    assert panel["words"] == "Not a working day"
    ctx = shifts_view.month_context(conn, NOW, e.id, 2026, 11, "2026-11-12")
    assert ctx["panel"]["state"] == "off"
    ctx = shifts_view.month_context(conn, NOW, e.id, 2026, 11, "2026-11-02")
    assert ctx["panel"]["scheduled"] is True and ctx["panel"]["note"] == ""
    assert ctx["back"].startswith(f"/shifts/month?e={e.id}&y=2026&m=11&day=2026-11-02")


def test_month_picks_the_schedule_of_the_person_using_this_device(conn):
    alex = people_repo.create_person(conn, "Alex", "A")
    sam = people_repo.create_person(conn, "Sam", "S")
    a = schedule(conn, title="Work", people_ids=[alex.id], owner_id=alex.id)
    s = schedule(conn, title="Clinic", people_ids=[sam.id], owner_id=sam.id)
    ctx = shifts_view.month_context(conn, NOW, me_id=sam.id)
    assert ctx["chosen"].id == s.id
    assert [x["label"] for x in ctx["schedules"]] == ["Alex: Work", "Sam: Clinic"]
    assert shifts_view.month_context(conn, NOW, a.id, me_id=sam.id)["chosen"].id == a.id
    assert shifts_view.month_context(conn, NOW)["chosen"].id == a.id


def test_week_start_of_the_calendar_does_not_move_the_pattern(conn):
    schedule(conn)
    cal.update_calendar(conn, {"week_start": "mon"})
    cells = {
        c["iso"]: c
        for row in shifts_view.month_context(conn, NOW, None, 2026, 11)["rows"]
        for c in row
    }
    assert cells["2026-11-02"]["state"] == "working"
    assert date(2026, 11, 2).weekday() == 0
