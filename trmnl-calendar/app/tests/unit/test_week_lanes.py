import io
from datetime import date

from PIL import Image

from trmnl_calendar.domain.models import Event, Section
from trmnl_calendar.domain.recurrence import occurrences
from trmnl_calendar.render import text as T
from trmnl_calendar.render.layouts import base, week_lanes
from trmnl_calendar.render.pipeline import RenderCache, render
from trmnl_calendar.render.scene import Rect, Text

SECS = (
    Section(1, "Events", 1),
    Section(2, "Work", 2),
    Section(3, "Lunch", 3),
    Section(4, "Dinner", 4),
)
SUN = date(2026, 10, 25)


def E(i, title, sec=1, start=SUN, **kw):
    return Event(id=i, title=title, section_id=sec, start_date=start, **kw)


def ctx(events, today=SUN, secs=SECS, **kw):
    occ = occurrences(
        events,
        None,
        *week_lanes.window(today, kw.get("week_start", "sun")),
        kw.get("week_start", "sun"),
    )
    return base.PlanContext(today=today, sections=secs, occurrences=tuple(occ), **kw)


def texts(scene):
    return [i for i in scene.items if isinstance(i, Text)]


def test_png_is_800x480_two_bit_and_uses_only_four_greys():
    r = render(
        ctx(
            [
                E(1, "Dentist", start_time="09:30"),
                E(2, "Trip", start=date(2026, 10, 27), end_date=date(2026, 10, 29)),
            ]
        )
    )
    im = Image.open(io.BytesIO(r.png))
    assert im.size == (800, 480)
    assert set(im.tobytes()) <= {0, 85, 170, 255}
    assert (r.width, r.height, r.bit_depth) == (800, 480, 2)


def test_same_inputs_same_hash_and_different_content_different_hash():
    a = render(ctx([E(1, "Dentist")]))
    assert a.hash == render(ctx([E(1, "Dentist")])).hash
    assert a.hash != render(ctx([E(1, "Dentist 2")])).hash


def test_text_rules_black_only_nothing_under_10px_inside_canvas():
    events = [
        E(i, f"Event number {i} with a long title", sec=1 + i % 4, start_time="10:00")
        for i in range(1, 25)
    ]
    scene, _ = week_lanes.plan(ctx(events))
    for t in texts(scene):
        assert t.gray == 0 and t.size >= 10
        w = T.measure(t.text, t.size, t.bold)
        left = t.x - (w if t.anchor == "r" else w / 2 if t.anchor == "m" else 0)
        assert 0 <= left and left + w <= scene.width, t


def test_title_is_week_range():
    scene, _ = week_lanes.plan(ctx([]))
    assert texts(scene)[0].text == "Oct 25 – 31, 2026"
    assert base.title_range(date(2026, 10, 26), date(2026, 11, 1)) == "Oct 26 – Nov 1, 2026"
    assert base.title_range(date(2026, 12, 27), date(2027, 1, 2)) == "Dec 27, 2026 – Jan 2, 2027"


def test_time_formats():
    f = base.fmt_time
    assert [
        f("09:30", "12"),
        f("10:00", "12"),
        f("12:00", "12"),
        f("00:15", "12"),
        f("17:30", "12"),
    ] == ["9:30a", "10a", "12p", "12:15a", "5:30p"]
    assert [f("09:30", "24"), f("17:30", "24")] == ["09:30", "17:30"]


def test_trip_is_one_bar_with_range_and_cut_arrows():
    scene, _ = week_lanes.plan(
        ctx([E(1, "Cabin", start=date(2026, 10, 27), end_date=date(2026, 10, 29))])
    )
    assert any(t.text == "Cabin - Oct 27–29" for t in texts(scene))
    scene, _ = week_lanes.plan(
        ctx([E(1, "Long trip", start=date(2026, 10, 29), end_date=date(2026, 11, 3))])
    )
    assert any(t.text == "Long trip ▸" for t in texts(scene))
    scene, _ = week_lanes.plan(
        ctx([E(1, "Long trip", start=date(2026, 10, 20), end_date=date(2026, 10, 26))])
    )
    assert any(t.text.startswith("◂ Long trip") for t in texts(scene))


def test_trip_spans_hidden_days_by_covering_visible_columns():
    weekdays = 0b0011111  # Mon..Fri
    ev = E(1, "Trip", start=date(2026, 10, 26), end_date=date(2026, 10, 30))
    scene, fit = week_lanes.plan(ctx([ev], weekdays=weekdays))
    bars = [i for i in scene.items if isinstance(i, Rect) and i.fill == 170 and i.h == 16]
    assert len(bars) == 1 and bars[0].w == 5 * ((800 - 24 - 58) // 5) - 6


def test_hidden_days_get_no_column_and_are_counted():
    ev = [E(1, "Sat thing", start=date(2026, 10, 31)), E(2, "Mon thing", start=date(2026, 10, 26))]
    scene, fit = week_lanes.plan(ctx(ev, weekdays=0b0011111))
    names = [t.text for t in texts(scene)]
    assert "Sat 31" not in names and "Mon 26" in names and "Mon thing" in names
    assert "+1 on hidden days" in names
    assert "hidden_day_events" in [f.code for f in fit.findings]


def test_today_hidden_highlights_next_visible_day():
    saturday = date(2026, 10, 31)
    scene, _ = week_lanes.plan(ctx([], today=saturday, weekdays=0b0011111))
    # Sat/Sun hidden and no visible day after Saturday this week: nothing is highlighted
    assert not [i for i in scene.items if isinstance(i, Rect) and i.fill == 170 and i.y == 44]
    scene, _ = week_lanes.plan(ctx([], today=SUN, weekdays=0b0011111))  # Sunday hidden -> Monday
    hl = [i for i in scene.items if isinstance(i, Rect) and i.fill == 170 and i.y == 44]
    assert len(hl) == 1 and hl[0].x == 12 + 58 + 2


def test_overflow_shows_plus_n_and_is_reported():
    events = [E(i, f"Thing {i}", start_time="08:00") for i in range(1, 15)]
    scene, fit = week_lanes.plan(ctx(events))
    assert any(t.text.startswith("+") and t.anchor == "r" for t in texts(scene))
    assert "events_overflow" in [f.code for f in fit.findings]


def test_many_categories_flag_small_lanes():
    secs = tuple(Section(i, f"Cat {i}", i) for i in range(1, 9))
    _, fit = week_lanes.plan(ctx([], secs=secs))
    assert "lane_too_small" not in [f.code for f in fit.findings]  # ~46 px each (docs/04 s8)
    heavy = tuple(Section(i, f"Cat {i}", i, weight=4 if i <= 4 else 1) for i in range(1, 9))
    _, fit = week_lanes.plan(ctx([], secs=heavy))
    assert "lane_too_small" in [f.code for f in fit.findings]


def test_banner_category_draws_strip_not_lane():
    secs = (Section(1, "Events", 1), Section(5, "Holidays", 2, placement="banner"))
    scene, _ = week_lanes.plan(ctx([E(1, "Halloween", sec=5, start=date(2026, 10, 31))], secs=secs))
    assert any(t.text == "Halloween" and t.bold and t.size == 11 for t in texts(scene))
    assert "Holidays" not in [t.text for t in texts(scene)]


def test_unproven_profile_is_reported():
    _, fit = week_lanes.plan(ctx([], width=600, height=448))
    assert "profile_unproven" in [f.code for f in fit.findings]


def test_week_start_monday_changes_window_and_title():
    scene, _ = week_lanes.plan(ctx([], week_start="mon"))
    assert texts(scene)[0].text == "Oct 19 – 25, 2026"


def test_render_cache_is_lru():
    c = RenderCache(size=2)
    c.put("a", b"1")
    c.put("b", b"2")
    c.get("a")
    c.put("c", b"3")
    assert c.get("a") == b"1" and c.get("b") is None and c.get("c") == b"3"


# ---- person badges (slice 2)

from trmnl_calendar.domain.models import Person  # noqa: E402

PEOPLE = {1: Person(1, "Alex", "A"), 2: Person(2, "Bea", "B"), 3: Person(3, "Chris", "Ch")}


def pctx(events, **kw):
    c = ctx(events, **kw)
    return base.PlanContext(**{**c.__dict__, "people": PEOPLE})


def boxes(scene):
    return [i for i in scene.items if isinstance(i, Rect) and i.h == 12 and i.outline == 0]


def test_badges_show_owner_first_and_all_white():
    ev = E(1, "Skating", people=(2, 1), owner_id=2)
    scene, _ = week_lanes.plan(pctx([ev]))
    bx = boxes(scene)
    assert [b.fill for b in bx] == [255, 255]  # all white; the owner is simply first
    assert bx[0].x < bx[1].x
    labels = [t.text for t in texts(scene) if t.bold and t.size == 10]
    assert labels == ["B", "A"]


def test_no_badges_for_everyone_events():
    scene, _ = week_lanes.plan(pctx([E(1, "Bin day")]))
    assert boxes(scene) == []


def test_two_letter_label_gets_a_wider_box_and_max_three_badges():
    scene, _ = week_lanes.plan(pctx([E(1, "X", people=(3,), owner_id=3)]))
    assert boxes(scene)[0].w == 18
    p4 = {**PEOPLE, 4: Person(4, "Dee", "D")}
    c = ctx([E(1, "X", people=(1, 2, 3, 4))])
    scene, _ = week_lanes.plan(base.PlanContext(**{**c.__dict__, "people": p4}))
    assert len(boxes(scene)) == 3


def test_badges_can_be_switched_off_and_unknown_people_are_skipped():
    ev = E(1, "Skating", people=(1, 99), owner_id=1)
    c = ctx([ev], options={"show_badges": False})
    scene, _ = week_lanes.plan(base.PlanContext(**{**c.__dict__, "people": PEOPLE}))
    assert boxes(scene) == []
    scene, _ = week_lanes.plan(pctx([ev]))
    assert len(boxes(scene)) == 1  # 99 has no person record


def test_badges_shorten_the_title_not_overlap_it():
    long = "A really quite long appointment title"
    plain, _ = week_lanes.plan(pctx([E(1, long)]))
    badged, _ = week_lanes.plan(pctx([E(1, long, people=(1, 2), owner_id=1)]))
    w_plain = max(T.measure(t.text, 12) for t in texts(plain) if t.size == 12 and not t.bold)
    w_badged = max(T.measure(t.text, 12) for t in texts(badged) if t.size == 12 and not t.bold)
    assert w_badged < w_plain
    b = boxes(badged)[0]
    cell_left = 12 + 58
    text_right = max(
        t.x + T.measure(t.text, 12) for t in texts(badged) if t.size == 12 and not t.bold
    )
    assert text_right < b.x and b.x > cell_left


def test_trip_bar_shows_badges_inside_the_bar():
    ev = E(
        1, "Cabin", start=date(2026, 10, 27), end_date=date(2026, 10, 29), people=(1, 2), owner_id=1
    )
    scene, _ = week_lanes.plan(pctx([ev]))
    bar = next(i for i in scene.items if isinstance(i, Rect) and i.h == 16)
    for b in boxes(scene):
        assert bar.x < b.x and b.x + b.w <= bar.x + bar.w
