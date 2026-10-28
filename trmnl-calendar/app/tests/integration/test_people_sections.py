import pytest

from trmnl_calendar.errors import InUse, NotFound, ValidationFailed
from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import calendar as cal
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories import people as ppl
from trmnl_calendar.storage.repositories import sections as sec


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def ev(conn, **kw):
    d = dict(title="Thing", section_id=1, start_date="2026-03-10")
    d.update(kw)
    return repo.create_event(conn, d)


# ---- people


def test_people_crud_and_limits(conn):
    a = ppl.create_person(conn, " Alex ", "a")
    b = ppl.create_person(conn, "Bea", "B")
    assert (a.name, a.label, a.position) == ("Alex", "a", 1) and b.position == 2
    assert [p.name for p in ppl.list_people(conn)] == ["Alex", "Bea"]
    with pytest.raises(ValidationFailed) as e:
        ppl.create_person(conn, "Other", "A")  # labels are unique ignoring case
    assert "label" in e.value.fields
    with pytest.raises(ValidationFailed):
        ppl.create_person(conn, "", "Z")
    with pytest.raises(ValidationFailed):
        ppl.create_person(conn, "Long label", "ABC")
    for i in range(8):
        ppl.create_person(conn, f"P{i}", f"{i}x")
    with pytest.raises(ValidationFailed) as e:
        ppl.create_person(conn, "One too many", "Q")
    assert "10" in e.value.fields["name"]


def test_person_update_retire_and_delete_rules(conn):
    a = ppl.create_person(conn, "Alex", "A")
    ev(conn, people_ids=[a.id])
    ppl.update_person(conn, a.id, {"name": "Alexa", "active": False})
    got = ppl.get_person(conn, a.id)
    assert got.name == "Alexa" and not got.active
    assert ppl.list_people(conn, include_inactive=False) == []
    with pytest.raises(InUse):
        ppl.delete_person(conn, a.id)
    b = ppl.create_person(conn, "Bea", "B")
    ppl.delete_person(conn, b.id)
    with pytest.raises(NotFound):
        ppl.get_person(conn, b.id)
    with pytest.raises(ValidationFailed):
        ppl.update_person(conn, a.id, {"nonsense": 1})


# ---- event people


def test_event_people_owner_rules(conn):
    a, b, c = (
        ppl.create_person(conn, n, label) for n, label in (("A", "A"), ("B", "B"), ("C", "C"))
    )
    solo = ev(conn, people_ids=[b.id])
    assert solo.people == (b.id,) and solo.owner_id == b.id  # one person is the owner
    many = ev(conn, people_ids=[a.id, b.id, c.id])
    assert many.owner_id is None and many.people == (a.id, b.id, c.id)
    owned = ev(conn, people_ids=[a.id, b.id], owner_id=b.id)
    assert owned.people == (b.id, a.id) and owned.owner_id == b.id  # owner first
    everyone = ev(conn)
    assert everyone.people == () and everyone.owner_id is None


def test_event_people_validation(conn):
    a = ppl.create_person(conn, "A", "A")
    with pytest.raises(ValidationFailed) as e:
        ev(conn, people_ids=[999])
    assert "people" in e.value.fields
    with pytest.raises(ValidationFailed) as e:
        ev(conn, people_ids=[a.id], owner_id=999)
    assert "owner_id" in e.value.fields
    ppl.update_person(conn, a.id, {"active": False})
    with pytest.raises(ValidationFailed):
        ev(conn, people_ids=[a.id])  # cannot newly attach a retired person


def test_patch_people_and_keep_when_not_mentioned(conn):
    a = ppl.create_person(conn, "A", "A")
    b = ppl.create_person(conn, "B", "B")
    e1 = ev(conn, people_ids=[a.id])
    kept = repo.update_event(conn, e1.id, {"title": "New"}, 1)
    assert kept.people == (a.id,) and kept.owner_id == a.id
    swapped = repo.update_event(conn, e1.id, {"people_ids": [b.id]}, 2)
    assert swapped.people == (b.id,) and swapped.owner_id == b.id
    everyone = repo.update_event(conn, e1.id, {"people_ids": []}, 3)
    assert everyone.people == () and everyone.owner_id is None
    both = repo.update_event(conn, e1.id, {"people_ids": [a.id, b.id], "owner_id": a.id}, 4)
    assert both.people == (a.id, b.id) and both.owner_id == a.id
    dropped = repo.update_event(conn, e1.id, {"people_ids": [b.id]}, 5)  # old owner removed
    assert dropped.owner_id == b.id


def test_retired_person_stays_on_old_events_and_edits(conn):
    a = ppl.create_person(conn, "A", "A")
    e1 = ev(conn, people_ids=[a.id])
    ppl.update_person(conn, a.id, {"active": False})
    out = repo.update_event(conn, e1.id, {"title": "Still fine"}, 1)
    assert out.people == (a.id,)
    again = repo.update_event(conn, e1.id, {"people_ids": [a.id]}, 2)
    assert again.people == (a.id,)


def test_listing_attaches_people(conn):
    a = ppl.create_person(conn, "A", "A")
    ev(conn, people_ids=[a.id])
    listed = repo.list_events(
        conn, __import__("datetime").date(2026, 3, 1), __import__("datetime").date(2026, 3, 31)
    )
    assert listed[0].people == (a.id,)


# ---- sections


def test_create_update_and_name_rules(conn):
    s = sec.create_section(conn, "Outings", weight=2, placement="banner")
    assert (s.name, s.weight, s.placement, s.position) == ("Outings", 2, "banner", 5)
    with pytest.raises(ValidationFailed) as e:
        sec.create_section(conn, "outings")
    assert "name" in e.value.fields
    out = sec.update_section(conn, s.id, {"weight": 4, "visible": False})
    assert out.weight == 4 and not out.visible and out.name == "Outings"
    with pytest.raises(ValidationFailed):
        sec.update_section(conn, s.id, {"weight": 9})
    with pytest.raises(ValidationFailed):
        sec.update_section(conn, s.id, {"placement": "sidebar"})


def test_limit_of_eight_categories(conn):
    for i in range(4):
        sec.create_section(conn, f"Extra {i}")
    with pytest.raises(ValidationFailed) as e:
        sec.create_section(conn, "Ninth")
    assert "8" in e.value.fields["name"]


def test_delete_moves_events_or_refuses(conn):
    e1 = ev(conn, section_id=2)
    with pytest.raises(InUse):
        sec.delete_section(conn, 2)
    assert sec.delete_section(conn, 2, move_to=1) == 1
    assert repo.get_event(conn, e1.id).section_id == 1
    assert 2 not in [s.id for s in cal.list_sections(conn)]
    assert sec.delete_section(conn, 3) == 0  # empty category needs no target


def test_cannot_delete_the_last_category(conn):
    for sid in (2, 3, 4):
        sec.delete_section(conn, sid)
    with pytest.raises(InUse):
        sec.delete_section(conn, 1)


def test_reorder_and_move(conn):
    sec.reorder_sections(conn, [4, 3, 2, 1])
    assert [s.id for s in cal.list_sections(conn)] == [4, 3, 2, 1]
    sec.move_section(conn, 1, -1)
    assert [s.id for s in cal.list_sections(conn)] == [4, 3, 1, 2]
    sec.move_section(conn, 4, -1)  # already first: no change
    assert [s.id for s in cal.list_sections(conn)][0] == 4
    with pytest.raises(ValidationFailed):
        sec.reorder_sections(conn, [1, 2, 3])


# ---- calendar settings


def test_calendar_settings(conn):
    out = cal.update_calendar(
        conn, {"timezone": "America/Toronto", "week_start": "mon", "time_format": "24"}
    )
    assert (out.timezone, out.week_start, out.time_format) == ("America/Toronto", "mon", "24")
    assert cal.get_calendar(conn).name == "Calendar"
    for bad in (
        {"timezone": "Mars/Base"},
        {"week_start": "tue"},
        {"time_format": "13"},
        {"name": ""},
    ):
        with pytest.raises(ValidationFailed):
            cal.update_calendar(conn, bad)
    with pytest.raises(ValidationFailed):
        cal.update_calendar(conn, {"colour": "red"})
