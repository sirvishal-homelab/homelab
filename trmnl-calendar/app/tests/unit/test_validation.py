from datetime import date

from trmnl_calendar.domain.models import Event
from trmnl_calendar.domain.repeat_text import repeat_text
from trmnl_calendar.domain.validation import validate_event


def good(**kw):
    d = dict(title="Dentist", section_id=1, start_date="2026-03-10")
    d.update(kw)
    return d


def test_minimal_valid():
    clean, err = validate_event(good())
    assert err == {} and clean["freq"] == "none" and clean["start_date"] == date(2026, 3, 10)


def test_required_fields():
    _, err = validate_event({})
    assert set(err) >= {"title", "section_id", "start_date"}


def test_title_limits_and_trim():
    assert validate_event(good(title="  hi  "))[0]["title"] == "hi"
    assert "title" in validate_event(good(title="x" * 121))[1]
    assert "title" in validate_event(good(title="   "))[1]


def test_end_date_rules():
    assert "end_date" in validate_event(good(end_date="2026-03-09"))[1]
    clean, err = validate_event(good(end_date="2026-03-10"))
    assert err == {} and clean["end_date"] is None  # same day is not a trip


def test_times():
    assert "start_time" in validate_event(good(start_time="9:5"))[1]
    assert "end_time" in validate_event(good(end_time="10:00"))[1]
    assert "end_time" in validate_event(good(start_time="10:00", end_time="10:00"))[1]
    assert (
        validate_event(good(start_time="22:00", end_time="06:00", end_date="2026-03-11"))[1] == {}
    )


def test_repeat_fields_cleared_when_not_repeating():
    clean, _ = validate_event(good(freq="none", interval=5, by_weekday=["MO"], until="2027-01-01"))
    assert clean["interval"] == 1 and clean["by_weekday"] == () and clean["until"] is None


def test_weekly_weekdays_sorted():
    clean, err = validate_event(good(freq="weekly", by_weekday=["FR", "MO"]))
    assert err == {} and clean["by_weekday"] == ("MO", "FR")


def test_monthly_weekday_needs_one_weekday_and_nth():
    _, err = validate_event(good(freq="monthly", month_mode="weekday", month_nth=1, by_weekday=[]))
    assert "by_weekday" in err
    _, err = validate_event(
        good(freq="monthly", month_mode="weekday", month_nth=5, by_weekday=["MO"])
    )
    assert "month_nth" in err
    clean, err = validate_event(
        good(freq="monthly", month_mode="weekday", month_nth="-1", by_weekday=["FR"])
    )
    assert err == {} and clean["month_nth"] == -1


def test_interval_bounds_and_until():
    assert "interval" in validate_event(good(freq="daily", interval=0))[1]
    assert "interval" in validate_event(good(freq="daily", interval=100))[1]
    assert "until" in validate_event(good(freq="daily", until="2026-03-01"))[1]


def e(**kw):
    base = dict(id=1, title="X", section_id=1, start_date=date(2026, 3, 5))
    base.update(kw)
    return Event(**base)


def test_repeat_text():
    assert repeat_text(e()) == "Does not repeat"
    assert repeat_text(e(freq="weekly", interval=2)) == "Repeats every 2 weeks on Thursday"
    assert (
        repeat_text(e(freq="weekly", by_weekday=("MO", "WE", "FR"), until=date(2026, 6, 30)))
        == "Repeats every week on Monday, Wednesday and Friday, until 30 June 2026"
    )
    assert (
        repeat_text(e(freq="monthly", month_mode="weekday", month_nth=-1, by_weekday=("FR",)))
        == "Repeats every month on the last Friday"
    )
    assert repeat_text(e(freq="yearly")) == "Repeats every year"
