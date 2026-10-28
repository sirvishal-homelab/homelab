from datetime import UTC, datetime

import pytest

from trmnl_calendar.errors import InUse, NotFound, ValidationFailed
from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import events as events_repo
from trmnl_calendar.storage.repositories import people as ppl
from trmnl_calendar.storage.repositories import sections as sec
from trmnl_calendar.storage.repositories import views as repo
from trmnl_calendar.web import forms, render_service, views

NOW = datetime(2026, 10, 28, 15, 0, tzinfo=UTC)


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def test_two_starter_views_and_active(conn):
    vs = repo.list_views(conn)
    assert [v.layout for v in vs] == ["week_lanes", "twoweek_list"]
    assert repo.active_view(conn).id == vs[0].id
    d = repo.get_display(conn)
    assert (d.width, d.height, d.bit_depth, d.model_name) == (800, 480, 2, "og_plus")


def test_create_from_template_and_name_rules(conn):
    v = repo.create_view(
        conn, "blank", {"name": "Work days", "weekdays": ["MO", "TU", "WE", "TH", "FR"]}
    )
    assert v.weekdays == 31 and v.layout == "week_lanes" and v.options["show_times"] is True
    with pytest.raises(ValidationFailed) as e:
        repo.create_view(conn, "blank", {"name": "work DAYS"})
    assert "name" in e.value.fields
    with pytest.raises(ValidationFailed):
        repo.create_view(conn, "blank", {"name": ""})


def test_validation_of_days_categories_people_options(conn):
    base = repo.template_values("week_lanes") | {"name": "X"}
    for bad, field in (
        ({"weekdays": []}, "weekdays"),
        ({"weekdays": ["XX"]}, "weekdays"),
        ({"categories": {"mode": "only", "ids": []}}, "categories"),
        ({"categories": {"mode": "only", "ids": [999]}}, "categories"),
        ({"people": {"mode": "only", "ids": [5]}}, "people"),
        ({"options": {"show_times": "yes"}}, "options"),
        ({"layout": "nope"}, "layout"),
        ({"surprise": 1}, "surprise"),
    ):
        with pytest.raises(ValidationFailed) as e:
            repo.build_view(conn, {**base, **bad})
        assert field in e.value.fields, bad


def test_changing_layout_resets_options(conn):
    v = repo.create_view(conn, "week_lanes", {"name": "A", "options": {"hide_empty_lanes": True}})
    assert v.options["hide_empty_lanes"] is True
    v2 = repo.update_view(conn, v.id, {"layout": "twoweek_list"})
    assert "hide_empty_lanes" not in v2.options and v2.layout == "twoweek_list"


def test_activate_delete_rules_and_reorder(conn):
    a, b = repo.list_views(conn)
    with pytest.raises(InUse):
        repo.delete_view(conn, a.id)  # active
    repo.activate_view(conn, b.id)
    repo.delete_view(conn, a.id)
    with pytest.raises(NotFound):
        repo.get_view(conn, a.id)
    c = repo.create_view(conn, "blank", {"name": "C"})
    repo.reorder_views(conn, [c.id, b.id])
    assert [v.id for v in repo.list_views(conn)] == [c.id, b.id]
    with pytest.raises(ValidationFailed):
        repo.reorder_views(conn, [c.id])


def test_render_uses_active_view_and_view_filters(conn):
    p1, p2 = ppl.create_person(conn, "A", "A"), ppl.create_person(conn, "B", "B")
    outings = sec.create_section(conn, "Outings")
    events_repo.create_event(
        conn, {"title": "Mine", "section_id": 1, "start_date": "2026-10-28", "people_ids": [p1.id]}
    )
    events_repo.create_event(
        conn, {"title": "Hers", "section_id": 1, "start_date": "2026-10-28", "people_ids": [p2.id]}
    )
    events_repo.create_event(
        conn, {"title": "Hike", "section_id": outings.id, "start_date": "2026-10-28"}
    )
    everything = render_service.render_current(conn, NOW)
    only_a = repo.create_view(
        conn, "blank", {"name": "A only", "people": {"mode": "only", "ids": [p1.id]}}
    )
    only_cat = repo.create_view(
        conn,
        "blank",
        {"name": "Outings", "categories": {"mode": "only", "ids": [outings.id]}},
    )
    texts = lambda v: [  # noqa: E731
        i.text for i in render_service.render_view(conn, v, NOW).scene.items if hasattr(i, "text")
    ]
    assert "Hers" not in texts(only_a) and "Mine" in texts(only_a)
    assert texts(only_cat).count("Hike") == 1 and "Mine" not in texts(only_cat)
    repo.activate_view(conn, only_a.id)
    assert render_service.render_current(conn, NOW).hash != everything.hash


def test_two_week_view_renders(conn):
    two = repo.list_views(conn)[1]
    r = render_service.render_view(conn, two, NOW)
    assert (r.width, r.height) == (800, 480)


def test_editor_and_list_pages_render(conn):
    ctx = views.views_list_context(conn)
    html = views.render_page("views.html", **ctx, flash="")
    assert "Two weeks - list" in html and "Make a new view" in html
    v = repo.get_view(conn, 1)
    ctx = views.view_editor_context(conn, NOW, None, repo.view_to_dict(v), v.id)
    html = views.render_page("view_edit.html", **ctx)
    assert 'name="weekday"' in html and "Everything fits" in html and "/render/" in html


def test_view_form_parsing(conn):
    f = forms.parse_body(
        b"name=Work&layout=twoweek_list&weekday=MO&weekday=TU&categories_mode=only&category=1"
        b"&people_mode=all&show_times=on"
    )
    d = forms.view_form_to_dict(f)
    assert d["weekdays"] == ["MO", "TU"] and d["categories"] == {"mode": "only", "ids": [1]}
    assert d["options"]["show_times"] is True and d["options"]["show_locations"] is False
    v = repo.create_view(conn, "blank", d)
    assert v.layout == "twoweek_list" and v.weekdays == 3
