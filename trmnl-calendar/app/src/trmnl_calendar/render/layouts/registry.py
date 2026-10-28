"""Layouts register by id; the editor lists whatever is here and builds its form from OPTIONS."""

from __future__ import annotations

from trmnl_calendar.render.layouts import twoweek_list, week_lanes
from trmnl_calendar.render.layouts.base import OPTION_DEFAULTS, OPTION_LABELS  # noqa: F401

LAYOUTS = {m.LAYOUT_ID: m for m in (week_lanes, twoweek_list)}


def layout_choices() -> list[tuple[str, str]]:
    return [(lid, m.LABEL) for lid, m in LAYOUTS.items()]


def option_keys(layout_id: str) -> tuple[str, ...]:
    return LAYOUTS[layout_id].OPTIONS


def default_options(layout_id: str) -> dict[str, bool]:
    return {k: OPTION_DEFAULTS[k] for k in option_keys(layout_id)}
