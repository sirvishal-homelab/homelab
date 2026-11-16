"""A long title is shortened before the flight status is (docs/21 section 4)."""

from dataclasses import replace
from datetime import date

import pytest

from trmnl_calendar.domain.models import Event, Section
from trmnl_calendar.domain.recurrence import occurrences
from trmnl_calendar.render import text as T
from trmnl_calendar.render.layouts import base, twoweek_list, week_lanes
from trmnl_calendar.render.scene import Text

SECS = (Section(1, "Events", 1), Section(2, "Work", 2))
MON = date(2026, 11, 16)


def plan(layout, title, word, **ev):
    e = Event(id=1, title=title, section_id=1, start_date=MON, start_time="11:25", **ev)
    occ = [
        replace(o, title=f"{o.title} ({word})", flight=word)
        for o in occurrences([e], None, *layout.window(MON, "sun"), "sun")
    ]
    ctx = base.PlanContext(today=MON, sections=SECS, occurrences=tuple(occ))
    scene, _ = layout.plan(ctx)
    return [i.text for i in scene.items if isinstance(i, Text)]


@pytest.mark.parametrize("layout", [week_lanes, twoweek_list])
def test_short_title_keeps_its_look_with_the_status_after_it(layout):
    lines = " ".join(plan(layout, "AC663", "Delayed"))
    assert "AC663" in lines and "(Delayed)" in lines


@pytest.mark.parametrize("layout", [week_lanes, twoweek_list])
def test_long_title_is_shortened_and_the_status_kept(layout):
    title = "Work trip to Montreal head office for the quarterly review AC663"
    texts = plan(layout, title, "Delayed")
    assert "(Delayed)" in texts
    assert any(t.endswith("…") for t in texts)


def test_trip_bar_keeps_the_status():
    texts = plan(
        week_lanes,
        "Montreal head office quarterly review AC663",
        "On time",
        end_date=date(2026, 11, 17),
    )
    assert any(t.endswith("(On time)") for t in texts)


def test_split_and_ellipsize_keep():
    assert T.split_tail("11:25 AC663 (Delayed)", "Delayed") == ("11:25 AC663", "(Delayed)")
    assert T.split_tail("Dentist", "") == ("Dentist", "")
    long = "A very long title indeed AC663 (Delayed)"
    text, cut = T.ellipsize_keep(long, "Delayed", 12, True, 90)
    assert cut and text.endswith("… (Delayed)")


@pytest.mark.parametrize("layout", [week_lanes, twoweek_list])
def test_delay_time_fits_without_a_false_shortened_warning(layout):
    e = Event(id=1, title="AC663", section_id=1, start_date=MON, start_time="11:25")
    occ = [
        replace(o, title=f"{o.title} (Delayed 1:25p)", flight="Delayed 1:25p")
        for o in occurrences([e], None, *layout.window(MON, "sun"), "sun")
    ]
    ctx = base.PlanContext(today=MON, sections=SECS, occurrences=tuple(occ))
    scene, fit = layout.plan(ctx)
    texts = [i.text for i in scene.items if isinstance(i, Text)]
    assert "(Delayed 1:25p)" in texts
    assert not [f for f in fit.findings if f.code == "titles_truncated"]
