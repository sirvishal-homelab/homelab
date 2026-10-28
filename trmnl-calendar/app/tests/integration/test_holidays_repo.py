import pytest

from trmnl_calendar.domain import holidays as H
from trmnl_calendar.errors import ValidationFailed
from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as events_repo
from trmnl_calendar.storage.repositories import holidays as repo
from trmnl_calendar.storage.repositories import sections as sec
from trmnl_calendar.web import views


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def default_items(year=2026):
    return [h for h in H.holidays_for("CA-NS", year) if h.group in H.CHECKED_BY_DEFAULT]


def test_import_creates_a_banner_category_and_events(conn):
    res = repo.import_holidays(conn, default_items())
    assert res.added == 10 and res.skipped == 0 and res.category_name == "Holidays"
    made = next(s for s in cal.list_sections(conn) if s.name == "Holidays")
    assert made.placement == "banner" and made.id == res.category_id
    events = events_repo.list_events(conn, *(__import__("datetime").date(2026, 1, 1),) * 2)
    assert [e.title for e in events] == ["New Year's Day"] and events[0].start_time is None


def test_rerun_skips_what_is_already_there_and_adds_new_years(conn):
    repo.import_holidays(conn, default_items(2026))
    again = repo.import_holidays(conn, default_items(2026) + default_items(2027))
    assert (again.added, again.skipped) == (10, 10)


def test_deleted_holiday_can_be_added_again(conn):
    repo.import_holidays(conn, default_items())
    ev = conn.execute("SELECT id, rev FROM events WHERE title = 'Labour Day'").fetchone()
    events_repo.soft_delete(conn, ev["id"], ev["rev"])
    again = repo.import_holidays(conn, default_items())
    assert (again.added, again.skipped) == (1, 9)


def test_existing_category_is_reused_and_choice_respected(conn):
    mine = sec.create_section(conn, "holidays", placement="lane")
    res = repo.import_holidays(conn, default_items())
    assert res.category_id == mine.id
    other = sec.create_section(conn, "Family")
    res = repo.import_holidays(conn, default_items()[:2], other.id)
    assert res.category_id == other.id and res.added == 2
    with pytest.raises(ValidationFailed):
        repo.import_holidays(conn, default_items(), 999)
    with pytest.raises(ValidationFailed):
        repo.import_holidays(conn, [])


def test_picker_page_model_marks_added_and_defaults(conn):
    repo.import_holidays(conn, default_items(2026)[:3])
    ctx = views.holidays_context(conn, "CA-NS", [2026], 2026)
    rows = {r["name"]: r for g in ctx["groups"] for r in g["rows"]}
    assert rows["New Year's Day"]["added"] and not rows["New Year's Day"]["checked"]
    assert rows["Canada Day"]["checked"] or rows["Canada Day"]["added"]
    assert not rows["Halloween"]["checked"]
    html = views.render_page("holidays.html", **ctx)
    assert "Nova Scotia, Canada" in html and "already added" in html
    assert views.render_page("holidays.html", **views.holidays_context(conn, "", [], 2026))
