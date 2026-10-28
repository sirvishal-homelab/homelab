from datetime import date

from trmnl_calendar.domain import holidays as H


def on(region, year):
    return {h.key: h for h in H.holidays_for(region, year)}


def test_easter_dates():
    assert H.easter(2025) == date(2025, 4, 20)
    assert H.easter(2026) == date(2026, 4, 5)
    assert H.easter(2027) == date(2027, 3, 28)


def test_nova_scotia_matches_the_provincial_chart_for_2025_and_2026():
    # novascotia.ca/lae/employmentrights/holidaychart.asp
    chart = {
        2025: {
            "new_year": date(2025, 1, 1),
            "heritage_day": date(2025, 2, 17),
            "good_friday": date(2025, 4, 18),
            "canada_day": date(2025, 7, 1),
            "labour_day": date(2025, 9, 1),
            "christmas": date(2025, 12, 25),
            "easter_sunday": date(2025, 4, 20),
            "thanksgiving": date(2025, 10, 13),
            "boxing_day": date(2025, 12, 26),
        },
        2026: {
            "new_year": date(2026, 1, 1),
            "heritage_day": date(2026, 2, 16),
            "good_friday": date(2026, 4, 3),
            "canada_day": date(2026, 7, 1),
            "labour_day": date(2026, 9, 7),
            "christmas": date(2026, 12, 25),
            "easter_sunday": date(2026, 4, 5),
            "thanksgiving": date(2026, 10, 12),
            "boxing_day": date(2026, 12, 26),
        },
    }
    for year, expected in chart.items():
        got = on("CA-NS", year)
        for key, d in expected.items():
            assert got[key].date == d, (year, key)


def test_nova_scotia_groups_follow_the_chart():
    g = {k: h.group for k, h in on("CA-NS", 2026).items()}
    assert {k for k, v in g.items() if v == "public"} == {
        "new_year",
        "heritage_day",
        "good_friday",
        "canada_day",
        "labour_day",
        "christmas",
    }
    assert g["thanksgiving"] == "closing" and g["remembrance_day"] == "closing"
    assert g["victoria_day"] == "common" and g["natal_day"] == "common"
    assert "heritage_day" not in on("CA", 2026)


def test_other_rules():
    got = on("CA-NS", 2026)
    assert got["victoria_day"].date == date(2026, 5, 18)
    assert on("CA-NS", 2027)["victoria_day"].date == date(2027, 5, 24)
    assert got["mothers_day"].date == date(2026, 5, 10)
    assert got["fathers_day"].date == date(2026, 6, 21)
    assert got["natal_day"].date == date(2026, 8, 3)
    assert got["dst_forward"].date == date(2026, 3, 8) and got["dst_back"].date == date(2026, 11, 1)


def test_sorted_and_unknown_region():
    days = [h.date for h in H.holidays_for("CA", 2026)]
    assert days == sorted(days)
    try:
        H.holidays_for("XX", 2026)
    except KeyError:
        pass
    else:
        raise AssertionError("expected KeyError")
