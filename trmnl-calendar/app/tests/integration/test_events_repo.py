from datetime import date

import pytest

from trmnl_calendar.errors import EditConflict, NotFound, ValidationFailed
from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import idempotency


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def new(conn, **kw):
    d = dict(title="Dentist", section_id=1, start_date="2026-03-10")
    d.update(kw)
    return repo.create_event(conn, d)


def test_seed(conn):
    assert [s.name for s in cal.list_sections(conn)] == ["Events", "Work", "Lunch", "Dinner"]
    c = cal.get_calendar(conn)
    assert c.timezone == "America/Halifax" and c.week_start == "sun" and c.time_format == "12"


def test_create_and_roundtrip(conn):
    ev = new(conn, start_time="09:30", freq="weekly", by_weekday=["TU", "TH"], until="2026-06-30")
    got = repo.get_event(conn, ev.id)
    assert got.rev == 1 and got.by_weekday == ("TU", "TH") and got.until == date(2026, 6, 30)
    assert got == ev


def test_create_rejects_bad_input_and_unknown_category(conn):
    with pytest.raises(ValidationFailed) as e:
        repo.create_event(conn, {})
    assert "title" in e.value.fields
    with pytest.raises(ValidationFailed) as e:
        new(conn, section_id=999)
    assert "section_id" in e.value.fields


def test_patch_changes_only_given_fields_and_bumps_rev(conn):
    ev = new(conn, location="Main St", start_time="09:00")
    out = repo.update_event(conn, ev.id, {"title": "Dentist (new)"}, 1)
    assert out.rev == 2 and out.title == "Dentist (new)"
    assert out.location == "Main St" and out.start_time == "09:00"


def test_patch_stale_rev_conflicts(conn):
    ev = new(conn)
    repo.update_event(conn, ev.id, {"title": "A"}, 1)
    with pytest.raises(EditConflict):
        repo.update_event(conn, ev.id, {"title": "B"}, 1)


def test_patch_validates_merged_result(conn):
    ev = new(conn, start_time="09:00", end_time="10:00")
    with pytest.raises(ValidationFailed):
        repo.update_event(conn, ev.id, {"end_time": "08:00"}, 1)
    with pytest.raises(ValidationFailed):
        repo.update_event(conn, ev.id, {"nonsense": 1}, 1)


def test_soft_delete_hides_and_restore_brings_back(conn):
    ev = new(conn)
    gone = repo.soft_delete(conn, ev.id, 1)
    with pytest.raises(NotFound):
        repo.get_event(conn, ev.id)
    assert [e.id for e in repo.list_deleted(conn)] == [gone.id]
    assert repo.list_events(conn, date(2026, 1, 1), date(2026, 12, 31)) == []
    back = repo.restore(conn, ev.id)
    assert back.id == ev.id and back.rev == 3
    assert repo.list_deleted(conn) == []


def test_old_deletes_are_purged_and_not_restorable(conn):
    ev = new(conn)
    repo.soft_delete(conn, ev.id, 1)
    conn.execute("UPDATE events SET deleted_at = '2020-01-01T00:00:00Z' WHERE id = ?", (ev.id,))
    with pytest.raises(NotFound):
        repo.restore(conn, ev.id)
    assert repo.purge_deleted(conn) == 1


def test_list_events_window_rules(conn):
    one = new(conn, title="once", start_date="2026-03-10")
    rep = new(conn, title="weekly", start_date="2026-01-01", freq="weekly")
    ended = new(conn, title="ended", start_date="2026-01-01", freq="daily", until="2026-01-31")
    trip = new(conn, title="trip", start_date="2026-02-25", end_date="2026-03-12")
    ids = {e.id for e in repo.list_events(conn, date(2026, 3, 8), date(2026, 3, 14))}
    assert ids == {one.id, rep.id, trip.id} and ended.id not in ids


def test_skips(conn):
    ev = new(conn, freq="weekly")
    repo.add_skip(conn, ev.id, "2026-03-17")
    repo.add_skip(conn, ev.id, "2026-03-17")  # repeat is harmless
    assert repo.skips_for(conn, [ev.id]) == {ev.id: {date(2026, 3, 17)}}
    repo.remove_skip(conn, ev.id, "2026-03-17")
    assert repo.skips_for(conn, [ev.id]) == {}
    with pytest.raises(ValidationFailed):
        repo.add_skip(conn, ev.id, "nonsense")


def test_idempotency_roundtrip_and_expiry(conn):
    assert idempotency.lookup(conn, "k1", "POST /events") is None
    idempotency.store(conn, "k1", "POST /events", 201, {"id": 5})
    assert idempotency.lookup(conn, "k1", "POST /events") == (201, {"id": 5})
    assert idempotency.lookup(conn, "k1", "POST /other") is None
    conn.execute("UPDATE idempotency_keys SET created_at = '2020-01-01T00:00:00Z'")
    assert idempotency.lookup(conn, "k1", "POST /events") is None
    assert idempotency.purge(conn) == 1


def test_database_constraints_back_up_validation(conn):
    import sqlite3

    with pytest.raises(sqlite3.IntegrityError):
        conn.execute(
            "INSERT INTO events (uid, calendar_id, section_id, title, start_date, freq, created_at,"
            " updated_at) VALUES ('u', 1, 1, 'x', '2026-01-01', 'hourly', 'n', 'n')"
        )
