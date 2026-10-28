from datetime import date

from trmnl_calendar.domain.models import Event
from trmnl_calendar.domain.validation import validate_event
from trmnl_calendar.web import forms
from trmnl_calendar.web.forms import event_to_values
from trmnl_calendar.web.serialize import event_to_api, patch_from_api


def f(**kw):
    base = {"title": "Dentist", "section_id": "1", "start_date": "2026-03-10", "repeat": "none"}
    base.update(kw)
    return {k: v if isinstance(v, list) else [v] for k, v in base.items()}


def data(**kw):
    return forms.form_to_event_data(f(**kw))


def test_parse_body_handles_unicode_blank_and_repeats():
    got = forms.parse_body(b"title=Caf%C3%A9+night&location=&by_weekday=MO&by_weekday=FR")
    assert got["title"] == ["Café night"] and got["location"] == [""]
    assert got["by_weekday"] == ["MO", "FR"]


def test_simple_repeat_choices():
    assert data(repeat="daily")["freq"] == "daily"
    d = data(repeat="biweekly")
    assert (d["freq"], d["interval"]) == ("weekly", 2)
    assert data()["freq"] == "none"


def test_blank_optionals_become_none_and_bad_section_is_none():
    d = data(end_date="", start_time="", section_id="abc")
    assert d["end_date"] is None and d["start_time"] is None and d["section_id"] is None


def test_custom_repeat_maps_all_fields_and_validates():
    d = data(
        repeat="custom",
        custom_freq="monthly",
        custom_interval="2",
        month_mode="weekday",
        month_nth="-1",
        by_weekday=["FR"],
        ends="until",
        until="2027-01-01",
    )
    clean, err = validate_event(d)
    assert err == {} and clean["freq"] == "monthly" and clean["month_nth"] == -1
    assert clean["interval"] == 2 and clean["until"] == date(2027, 1, 1)


def test_empty_title_error_is_plain_language():
    _, err = validate_event(data(title="  "))
    assert err["title"] == "Enter a short description"


def ev(**kw):
    base = dict(id=1, title="X", section_id=2, start_date=date(2026, 3, 5))  # a Thursday
    base.update(kw)
    return Event(**base)


def test_repeat_choice_is_inferred_for_editing():
    v = forms.event_to_values
    assert v(ev())["repeat"] == "none"
    assert v(ev(freq="daily"))["repeat"] == "daily"
    assert v(ev(freq="weekly"))["repeat"] == "weekly"
    assert v(ev(freq="weekly", by_weekday=("TH",)))["repeat"] == "weekly"
    assert v(ev(freq="weekly", interval=2))["repeat"] == "biweekly"
    assert v(ev(freq="weekly", by_weekday=("MO", "TH")))["repeat"] == "custom"
    assert v(ev(freq="monthly", until=date(2027, 1, 1)))["repeat"] == "monthly"
    assert (
        v(ev(freq="monthly", month_mode="weekday", month_nth=1, by_weekday=("MO",)))["repeat"]
        == "custom"
    )


def test_form_roundtrip_preserves_an_event():
    original = ev(
        freq="weekly", by_weekday=("MO", "TH"), interval=3, until=date(2026, 12, 1),
        end_date=date(2026, 3, 6), start_time="09:00", end_time="10:00", location="Arena",
    )  # fmt: skip
    values = forms.event_to_values(original)
    submitted = {k: v if isinstance(v, list) else [v] for k, v in values.items()}
    clean, err = validate_event(forms.form_to_event_data(submitted))
    assert err == {}
    assert (clean["freq"], clean["interval"], clean["by_weekday"]) == ("weekly", 3, ("MO", "TH"))
    assert clean["until"] == date(2026, 12, 1) and clean["end_date"] == date(2026, 3, 6)


def test_api_serialisation_and_patch_mapping():
    out = event_to_api(ev(freq="weekly", interval=2), {"updated_at": "t", "deleted_at": None})
    assert out["category_id"] == 2 and out["repeat"]["interval"] == 2
    assert out["repeat_text"] == "Repeats every 2 weeks on Thursday" and out["updated_at"] == "t"
    assert out["people"] == {"ids": [], "owner_id": None} and out["updated_by"] is None
    patch = patch_from_api(
        {"title": "T", "category_id": 3, "repeat": {"freq": "daily", "weekdays": []}}
    )
    assert patch == {"title": "T", "section_id": 3, "freq": "daily", "by_weekday": []}


def test_people_in_forms_and_api():
    d = data(people=["2", "1", "x"], owner_id="2")
    assert d["people_ids"] == [2, 1] and d["owner_id"] == 2
    assert data()["people_ids"] == [] and data()["owner_id"] is None
    e = ev(people=(2, 1), owner_id=2)
    v = forms.event_to_values(e)
    assert v["people"] == ["2", "1"] and v["owner_id"] == "2"
    assert forms.event_to_values(None, default_people=(3,))["people"] == ["3"]
    echoed = forms.values_from_form({"people": ["1", "2"], "owner_id": ["1"]})
    assert echoed["people"] == ["1", "2"] and echoed["owner_id"] == "1"
    from trmnl_calendar.domain.models import Person

    out = event_to_api(e, {"updated_by_person_id": 2}, {2: Person(2, "Bea", "B")})
    assert out["people"] == {"ids": [2, 1], "owner_id": 2}
    assert out["updated_by"] == {"id": 2, "label": "B"}
    patch = patch_from_api({"people": {"ids": [1], "owner_id": 1}})
    assert patch == {"people_ids": [1], "owner_id": 1}


def test_ends_choice_decides_between_until_and_count():
    base = dict(repeat="custom", custom_freq="weekly", until="2027-01-01", count="5")
    d = data(ends="count", **base)
    assert d["count"] == "5" and d["until"] is None
    d = data(ends="until", **base)
    assert d["until"] == "2027-01-01" and d["count"] is None
    d = data(ends="never", **base)
    assert d["until"] is None and d["count"] is None
    clean, err = validate_event(data(ends="count", **base))
    assert err == {} and clean["count"] == 5


def test_count_form_values_round_trip():
    ev = Event(
        id=1, title="Class", section_id=1, start_date=date(2026, 11, 3), freq="weekly", count=6
    )
    v = event_to_values(ev)
    assert v["ends"] == "count" and v["count"] == "6" and v["repeat"] == "weekly"


def test_ends_applies_to_preset_repeats_too():
    d = data(repeat="weekly", ends="count", count="5")
    assert d["freq"] == "weekly" and d["count"] == "5"
    d = data(repeat="daily", ends="until", until="2027-01-01")
    assert d["until"] == "2027-01-01"
    d = data(repeat="none", ends="count", count="5")
    assert d["count"] is None
