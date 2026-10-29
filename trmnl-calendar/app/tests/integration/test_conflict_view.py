import pytest

from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import people as people_repo
from trmnl_calendar.web import forms, views


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def test_conflict_lists_only_what_differs_and_who_changed_it(conn):
    p = people_repo.create_person(conn, "Madelyn", "M")
    ev = repo.create_event(conn, dict(title="Dentist", section_id=1, start_date="2026-03-10"))
    ev = repo.update_event(
        conn, ev.id, {"title": "Dentist 2", "start_time": "09:30"}, ev.rev, person_id=p.id
    )
    yours = forms.event_to_values(ev) | {"title": "Orthodontist", "start_time": "09:30"}
    c = views.conflict_context(conn, ev, yours)
    assert c["who"] == "Madelyn" and c["when"]
    assert c["diffs"] == [{"label": "Title", "saved": "Dentist 2", "yours": "Orthodontist"}]


def test_conflict_notes_repeat_changes(conn):
    ev = repo.create_event(
        conn, dict(title="Class", section_id=1, start_date="2026-03-10", freq="weekly")
    )
    yours = forms.event_to_values(ev) | {"repeat": "none"}
    diffs = views.conflict_context(conn, ev, yours)["diffs"]
    assert [d["label"] for d in diffs] == ["Repeats"]
