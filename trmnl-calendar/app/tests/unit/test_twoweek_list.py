from datetime import date

from trmnl_calendar.domain.models import Event, Section
from trmnl_calendar.domain.recurrence import occurrences
from trmnl_calendar.render.layouts import base, twoweek_list
from trmnl_calendar.render.pipeline import render
from trmnl_calendar.render.scene import Glyph, Rect, Text

SECS = (
    Section(1, "Events", 1),
    Section(2, "Work", 2),
    Section(5, "Holidays", 5, placement="banner"),
)
SUN = date(2026, 10, 25)
TODAY = date(2026, 10, 28)


def ctx(events, sections=SECS, **kw):
    first, last = twoweek_list.window(TODAY, "sun")
    occ = occurrences(events, {}, first, last, "sun")
    return base.PlanContext(today=TODAY, sections=sections, occurrences=tuple(occ), **kw)


def E(i, title, sec=1, start=SUN, **kw):
    return Event(id=i, title=title, section_id=sec, start_date=start, **kw)


def texts(scene):
    return [i.text for i in scene.items if isinstance(i, Text)]


def test_window_is_two_weeks_from_week_start():
    assert twoweek_list.window(TODAY, "sun") == (SUN, date(2026, 11, 7))
    assert twoweek_list.window(TODAY, "mon") == (date(2026, 10, 26), date(2026, 11, 8))


def test_title_and_events_shown_with_symbols():
    scene, fit = twoweek_list.plan(ctx([E(1, "Dentist", start_time="09:30"), E(2, "Review", 2)]))
    t = texts(scene)
    assert "Oct 25 – Nov 7, 2026" in t and "9:30a Dentist" in t and "Review" in t
    assert (
        sum(isinstance(i, Glyph) for i in scene.items) >= 4
    )  # two in the lists, two in the legend


def test_banner_category_draws_a_grey_strip_not_a_list_item():
    scene, _ = twoweek_list.plan(ctx([E(1, "Halloween", 5, date(2026, 10, 31))]))
    assert "Halloween" in texts(scene)
    assert any(isinstance(i, Rect) and i.fill == 170 and i.h == 14 for i in scene.items)


def test_banner_option_off_hides_banner_events():
    scene, _ = twoweek_list.plan(
        ctx([E(1, "Halloween", 5, date(2026, 10, 31))], options={"show_banners": False})
    )
    assert "Halloween" not in texts(scene)


def test_trip_bar_has_range_and_breaks_at_row_end():
    trip = E(1, "Visit Mom", start=date(2026, 10, 30), end_date=date(2026, 11, 2))
    scene, _ = twoweek_list.plan(ctx([trip]))
    t = texts(scene)
    assert any(x.endswith("▸") for x in t) and any(x.startswith("◂") for x in t)


def test_weekdays_mask_removes_columns_and_counts_hidden():
    work = 0b0011111  # Mon-Fri
    scene, fit = twoweek_list.plan(ctx([E(1, "Brunch", start=date(2026, 10, 31))], weekdays=work))
    assert "Sat" not in texts(scene)
    assert "+1 on hidden days" in texts(scene)
    assert any(f.code == "hidden_day_events" for f in fit.findings)


def test_overflow_reports_and_never_draws_outside_the_page():
    many = [E(i, f"Event number {i}", start=SUN) for i in range(1, 30)]
    scene, fit = twoweek_list.plan(ctx(many))
    assert any(f.code == "events_overflow" for f in fit.findings)
    assert any(t.startswith("+") for t in texts(scene))


def test_no_text_under_10px_and_renders_png():
    scene, _ = twoweek_list.plan(ctx([E(1, "Dentist")]))
    assert all(i.size >= 10 for i in scene.items if isinstance(i, Text))
    result = render(ctx([E(1, "Dentist")]), "twoweek_list")
    assert result.png.startswith(b"\x89PNG") and (result.width, result.height) == (800, 480)


def test_same_input_same_hash():
    a = render(ctx([E(1, "Dentist")]), "twoweek_list").hash
    b = render(ctx([E(1, "Dentist")]), "twoweek_list").hash
    assert a == b


def test_week_lanes_options_hide_empty_lanes_and_locations():
    from trmnl_calendar.render.layouts import week_lanes

    secs = (Section(1, "Events", 1), Section(2, "Work", 2))
    events = [E(1, "Dentist", location="Clinic", start_time="09:00")]
    first, last = week_lanes.window(TODAY, "sun")
    occ = tuple(occurrences(events, {}, first, last, "sun"))
    full = week_lanes.plan(base.PlanContext(today=TODAY, sections=secs, occurrences=occ))[0]
    slim = week_lanes.plan(
        base.PlanContext(
            today=TODAY,
            sections=secs,
            occurrences=occ,
            options={"hide_empty_lanes": True, "show_locations": True, "show_times": False},
        )
    )[0]
    assert "Work" in texts(full) and "Work" not in texts(slim)
    assert any("Dentist @ Clinic" in t for t in texts(slim))


# ---- person badges (shared with Week - lanes)

from trmnl_calendar.domain.models import Person  # noqa: E402
from trmnl_calendar.render import text as T  # noqa: E402

PEOPLE = {1: Person(1, "Alex", "A"), 2: Person(2, "Bea", "B")}


def pctx(events, **kw):
    c = ctx(events, **kw)
    return base.PlanContext(**{**c.__dict__, "people": PEOPLE})


def tx(scene):
    return [i for i in scene.items if isinstance(i, Text)]


def boxes(scene):
    return [i for i in scene.items if isinstance(i, Rect) and i.h == 12 and i.outline == 0]


def test_badges_are_on_by_default_owner_first_and_all_white():
    scene, _ = twoweek_list.plan(pctx([E(1, "Skating", people=(2, 1), owner_id=2)]))
    bx = boxes(scene)
    assert [b.fill for b in bx] == [255, 255] and bx[0].x < bx[1].x
    assert [t.text for t in tx(scene) if t.bold and t.size == 10] == ["B", "A"]


def test_everyone_events_have_no_badge_and_the_option_turns_them_off():
    scene, _ = twoweek_list.plan(pctx([E(1, "Garbage day")]))
    assert boxes(scene) == []
    ev = E(1, "Skating", people=(1,), owner_id=1)
    scene, _ = twoweek_list.plan(pctx([ev], options={"show_badges": False}))
    assert boxes(scene) == []
    assert "show_badges" in twoweek_list.OPTIONS


def test_badges_sit_on_the_first_line_and_long_titles_wrap_instead_of_being_cut():
    title = "Hot lunch with the whole class today"
    plain, _ = twoweek_list.plan(pctx([E(1, title)]))
    badged, _ = twoweek_list.plan(pctx([E(1, title, people=(1, 2), owner_id=1)]))
    shown = [t for t in badged.items if isinstance(t, Text) and not t.bold and t.size == 11]
    assert " ".join(t.text for t in shown).replace("…", "").split()[:3] == title.split()[:3]
    assert len(shown) >= len(
        [
            t
            for t in plain.items
            if isinstance(t, Text) and t.size == 11 and not t.bold and t.text in title
        ]
    )
    first = shown[0]
    assert first.x + T.measure(first.text, 11) < boxes(badged)[0].x  # never overlaps the box


def test_wrap_helper_uses_full_width_after_the_first_line():
    lines, cut = twoweek_list._wrap_with_badges("aaa bbb ccc ddd", 200, 0, 3)
    assert lines == ["aaa bbb ccc ddd"] and not cut
    lines, cut = twoweek_list._wrap_with_badges("alpha beta gamma delta epsilon", 80, 40, 3)
    assert len(lines) >= 2 and T.measure(lines[0], 11) <= 40
