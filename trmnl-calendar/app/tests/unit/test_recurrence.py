from datetime import date

from trmnl_calendar.domain.models import Event
from trmnl_calendar.domain.recurrence import nth_weekday, occurrences, week_floor


def ev(**kw):
    base = dict(id=1, title="X", section_id=1, start_date=date(2026, 1, 1))
    base.update(kw)
    return Event(**base)


def days(events, a, b, skips=None, week_start="sun"):
    return [o.date for o in occurrences(events, skips, a, b, week_start)]


def test_single_event_in_and_out_of_window():
    e = ev(start_date=date(2026, 3, 10))
    assert days([e], date(2026, 3, 1), date(2026, 3, 31)) == [date(2026, 3, 10)]
    assert days([e], date(2026, 3, 11), date(2026, 3, 31)) == []


def test_daily_interval_and_window_alignment():
    e = ev(freq="daily", interval=3, start_date=date(2026, 3, 1))
    assert days([e], date(2026, 3, 5), date(2026, 3, 13)) == [
        date(2026, 3, 7),
        date(2026, 3, 10),
        date(2026, 3, 13),
    ]


def test_weekly_default_weekday_is_start_date_weekday():
    e = ev(freq="weekly", start_date=date(2026, 3, 4))  # Wednesday
    assert days([e], date(2026, 3, 1), date(2026, 3, 22)) == [
        date(2026, 3, 4),
        date(2026, 3, 11),
        date(2026, 3, 18),
    ]


def test_weekly_multiple_weekdays_start_midweek():
    e = ev(freq="weekly", by_weekday=("MO", "FR"), start_date=date(2026, 3, 4))  # Wed
    assert days([e], date(2026, 3, 1), date(2026, 3, 16)) == [
        date(2026, 3, 6),
        date(2026, 3, 9),
        date(2026, 3, 13),
        date(2026, 3, 16),
    ]


def test_biweekly_counts_from_start_week():
    e = ev(freq="weekly", interval=2, start_date=date(2026, 3, 5))  # Thursday
    assert days([e], date(2026, 3, 1), date(2026, 4, 12)) == [
        date(2026, 3, 5),
        date(2026, 3, 19),
        date(2026, 4, 2),
    ]


def test_biweekly_week_start_matters_for_multi_weekday():
    # start Sunday Mar 1 2026. Sunday-start weeks: Sun Mar 1 begins week 0.
    e = ev(freq="weekly", interval=2, by_weekday=("SU", "MO"), start_date=date(2026, 3, 1))
    sun = days([e], date(2026, 3, 1), date(2026, 3, 20), week_start="sun")
    mon = days([e], date(2026, 3, 1), date(2026, 3, 20), week_start="mon")
    assert sun == [date(2026, 3, 1), date(2026, 3, 2), date(2026, 3, 15), date(2026, 3, 16)]
    assert mon == [date(2026, 3, 1), date(2026, 3, 9), date(2026, 3, 15)]


def test_monthly_by_date_clamps_to_month_end():
    e = ev(freq="monthly", start_date=date(2026, 1, 31))
    assert days([e], date(2026, 1, 1), date(2026, 5, 31)) == [
        date(2026, 1, 31),
        date(2026, 2, 28),
        date(2026, 3, 31),
        date(2026, 4, 30),
        date(2026, 5, 31),
    ]


def test_monthly_leap_february():
    e = ev(freq="monthly", start_date=date(2027, 12, 30))
    assert date(2028, 2, 29) in days([e], date(2028, 2, 1), date(2028, 2, 29))


def test_monthly_every_two_months():
    e = ev(freq="monthly", interval=2, start_date=date(2026, 1, 15))
    assert days([e], date(2026, 1, 1), date(2026, 7, 31)) == [
        date(2026, 1, 15),
        date(2026, 3, 15),
        date(2026, 5, 15),
        date(2026, 7, 15),
    ]


def test_yearly_feb29_clamps_in_common_years():
    e = ev(freq="yearly", start_date=date(2024, 2, 29))
    got = []
    for y in range(2024, 2029):
        got += days([e], date(y, 1, 1), date(y, 12, 31))
    assert got == [
        date(2024, 2, 29),
        date(2025, 2, 28),
        date(2026, 2, 28),
        date(2027, 2, 28),
        date(2028, 2, 29),
    ]


def test_nth_weekday_helper():
    assert nth_weekday(2026, 11, 0, 1) == date(2026, 11, 2)  # first Monday
    assert nth_weekday(2026, 11, 4, -1) == date(2026, 11, 27)  # last Friday
    assert nth_weekday(2026, 2, 0, 4) == date(2026, 2, 23)


def test_monthly_first_monday():
    e = ev(
        freq="monthly",
        month_mode="weekday",
        month_nth=1,
        by_weekday=("MO",),
        start_date=date(2026, 11, 2),
    )
    assert days([e], date(2026, 11, 1), date(2027, 1, 31)) == [
        date(2026, 11, 2),
        date(2026, 12, 7),
        date(2027, 1, 4),
    ]


def test_monthly_last_friday_start_before_pattern_waits_for_next_match():
    # started on Nov 20; Nov's last Friday is the 27th, which is on/after the start, so it counts
    e = ev(
        freq="monthly",
        month_mode="weekday",
        month_nth=-1,
        by_weekday=("FR",),
        start_date=date(2026, 11, 20),
    )
    assert days([e], date(2026, 11, 1), date(2027, 1, 31)) == [
        date(2026, 11, 27),
        date(2026, 12, 25),
        date(2027, 1, 29),
    ]


def test_monthly_weekday_pattern_date_before_start_is_skipped():
    e = ev(
        freq="monthly",
        month_mode="weekday",
        month_nth=1,
        by_weekday=("MO",),
        start_date=date(2026, 11, 10),
    )  # Nov's first Monday (2nd) is before the start
    assert days([e], date(2026, 11, 1), date(2026, 12, 31)) == [date(2026, 12, 7)]


def test_until_is_inclusive_on_start_date():
    e = ev(freq="daily", start_date=date(2026, 3, 1), until=date(2026, 3, 3))
    assert days([e], date(2026, 3, 1), date(2026, 3, 10)) == [
        date(2026, 3, 1),
        date(2026, 3, 2),
        date(2026, 3, 3),
    ]


def test_skip_removes_one_occurrence():
    e = ev(freq="weekly", start_date=date(2026, 3, 3))
    got = days([e], date(2026, 3, 1), date(2026, 3, 24), skips={1: {date(2026, 3, 10)}})
    assert got == [date(2026, 3, 3), date(2026, 3, 17), date(2026, 3, 24)]


def test_trip_yields_one_occurrence_per_day_with_flags_and_times():
    e = ev(
        start_date=date(2026, 11, 3),
        end_date=date(2026, 11, 6),
        start_time="09:00",
        end_time="17:30",
    )
    occ = occurrences([e], None, date(2026, 11, 1), date(2026, 11, 30))
    assert [o.date.day for o in occ] == [3, 4, 5, 6]
    assert [o.day_number for o in occ] == [1, 2, 3, 4] and occ[0].day_count == 4
    assert occ[0].is_first and occ[-1].is_last
    assert (occ[0].start_time, occ[0].end_time) == ("09:00", None)
    assert (occ[-1].start_time, occ[-1].end_time) == (None, "17:30")


def test_trip_started_before_window_still_appears():
    e = ev(start_date=date(2026, 11, 3), end_date=date(2026, 11, 9))
    occ = occurrences([e], None, date(2026, 11, 8), date(2026, 11, 14))
    assert [o.date.day for o in occ] == [8, 9] and occ[0].day_number == 6


def test_repeating_trip_each_occurrence_spans():
    e = ev(freq="weekly", start_date=date(2026, 3, 2), end_date=date(2026, 3, 3))
    assert days([e], date(2026, 3, 1), date(2026, 3, 15)) == [
        date(2026, 3, 2),
        date(2026, 3, 3),
        date(2026, 3, 9),
        date(2026, 3, 10),
    ]


def test_ordering_all_day_then_time_then_title():
    a = ev(id=1, title="B", start_date=date(2026, 3, 1), start_time="10:00")
    b = ev(id=2, title="A", start_date=date(2026, 3, 1))
    c = ev(id=3, title="C", start_date=date(2026, 3, 1), start_time="09:00")
    assert [
        o.event_id for o in occurrences([a, b, c], None, date(2026, 3, 1), date(2026, 3, 1))
    ] == [2, 3, 1]


def test_week_floor():
    assert week_floor(date(2026, 3, 4), "sun") == date(2026, 3, 1)
    assert week_floor(date(2026, 3, 1), "sun") == date(2026, 3, 1)
    assert week_floor(date(2026, 3, 4), "mon") == date(2026, 3, 2)
    assert week_floor(date(2026, 3, 1), "mon") == date(2026, 2, 23)


def test_window_limit():
    import pytest

    with pytest.raises(ValueError):
        occurrences([], None, date(2026, 1, 1), date(2028, 1, 1))


# ---- repeat count (the first occurrence counts; skips still count)


def test_count_limits_a_daily_series():
    e = ev(start_date=date(2026, 3, 2), freq="daily", count=4)
    assert days([e], date(2026, 3, 1), date(2026, 3, 31)) == [
        date(2026, 3, d) for d in (2, 3, 4, 5)
    ]


def test_count_with_several_weekdays_counts_each_day():
    e = ev(start_date=date(2026, 3, 2), freq="weekly", by_weekday=("MO", "WE"), count=5)
    got = days([e], date(2026, 3, 1), date(2026, 4, 30))
    assert got == [date(2026, 3, d) for d in (2, 4, 9, 11, 16)]


def test_count_survives_a_window_that_starts_later():
    e = ev(start_date=date(2026, 3, 2), freq="weekly", count=6)  # Mondays: Mar 2 .. Apr 6
    got = days([e], date(2026, 3, 20), date(2026, 5, 31))
    assert got == [date(2026, 3, 23), date(2026, 3, 30), date(2026, 4, 6)]


def test_skipped_dates_still_count():
    e = ev(start_date=date(2026, 3, 2), freq="weekly", count=3)
    got = days([e], date(2026, 3, 1), date(2026, 5, 31), {1: {date(2026, 3, 9)}})
    assert got == [date(2026, 3, 2), date(2026, 3, 16)]


def test_count_monthly_clamp_and_yearly_and_trips():
    monthly = ev(start_date=date(2026, 1, 31), freq="monthly", count=3)
    assert days([monthly], date(2026, 1, 1), date(2026, 12, 31)) == [
        date(2026, 1, 31),
        date(2026, 2, 28),
        date(2026, 3, 31),
    ]
    yearly = ev(id=2, start_date=date(2026, 6, 1), freq="yearly", count=2)
    assert days([yearly], date(2026, 5, 1), date(2026, 12, 31)) == [date(2026, 6, 1)]
    assert days([yearly], date(2027, 5, 1), date(2027, 12, 31)) == [date(2027, 6, 1)]
    assert days([yearly], date(2028, 5, 1), date(2028, 12, 31)) == []  # the series has ended
    trip = ev(id=3, start_date=date(2026, 7, 1), end_date=date(2026, 7, 3), freq="weekly", count=2)
    assert len(days([trip], date(2026, 7, 1), date(2026, 8, 31))) == 6  # 2 x 3 days


def test_count_on_a_non_repeating_event_is_ignored():
    e = ev(start_date=date(2026, 3, 2), count=3)
    assert len(days([e], date(2026, 3, 1), date(2026, 3, 31))) == 1


# ---- rotating weeks (work schedule)

A_WEEK = ("MO", "TU", "FR", "SA", "SU")
B_WEEK = ("WE", "TH")


def rota(**kw):
    base = dict(
        start_date=date(2026, 10, 26),  # a Monday: week A is the week of Oct 26
        freq="weekly",
        interval=2,
        rotation=(A_WEEK, B_WEEK),
    )
    base.update(kw)
    return ev(**base)


def test_rotation_alternates_two_different_weeks():
    got = days([rota()], date(2026, 10, 26), date(2026, 11, 22))
    want = [(10, 26), (10, 27), (10, 30), (10, 31), (11, 1), (11, 4), (11, 5)]
    want += [(11, 9), (11, 10), (11, 13), (11, 14), (11, 15), (11, 18), (11, 19)]
    assert got == [date(2026, m, d) for m, d in want]


def test_rotation_ignores_the_calendars_week_start():
    e = rota()
    a = days([e], date(2026, 10, 20), date(2026, 12, 31), week_start="sun")
    b = days([e], date(2026, 10, 20), date(2026, 12, 31), week_start="mon")
    assert a == b


def test_rotation_starting_mid_week_skips_earlier_days():
    e = rota(start_date=date(2026, 10, 30))  # Friday of week A
    got = days([e], date(2026, 10, 1), date(2026, 11, 8))
    assert got[0] == date(2026, 10, 30) and date(2026, 10, 27) not in got


def test_rotation_window_can_start_late_and_skips_apply():
    e = rota()
    skips = {1: {date(2026, 11, 14)}}  # the Saturday off
    got = days([e], date(2026, 11, 9), date(2026, 11, 15), skips)
    assert got == [date(2026, 11, d) for d in (9, 10, 13, 15)]


def test_rotation_count_and_until():
    assert len(days([rota(count=3)], date(2026, 10, 1), date(2026, 12, 31))) == 3
    e = rota(until=date(2026, 11, 1))
    assert days([e], date(2026, 10, 1), date(2026, 12, 31))[-1] == date(2026, 11, 1)


def test_day_note_goes_into_the_title():
    e = rota(title="Work")
    notes = {1: {date(2026, 10, 30): "out"}}
    occ = occurrences([e], None, date(2026, 10, 26), date(2026, 11, 1), "sun", notes)
    titles = {o.date.day: o.title for o in occ}
    assert titles[30] == "Work (out)" and titles[27] == "Work"
    assert all(o.scheduled for o in occ) and next(o for o in occ if o.date.day == 30).note == "out"
