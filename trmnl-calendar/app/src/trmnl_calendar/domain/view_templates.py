"""Starting values for new views. Templates live in code, not the database (docs/03, docs/17).

A template is copied into a new view; editing the view never changes the template.
"""

from __future__ import annotations

TEMPLATES: dict[str, dict] = {
    "week_lanes": {
        "label": "Week - lanes",
        "help": "One week. Categories run down the page, days run across.",
        "values": {"name": "Week - lanes", "layout": "week_lanes", "weekdays": 127},
    },
    "twoweek_list": {
        "label": "Two weeks - list",
        "help": "Two weeks at a glance. Each day is a short list with category symbols.",
        "values": {"name": "Two weeks - list", "layout": "twoweek_list", "weekdays": 127},
    },
    "blank": {
        "label": "Blank",
        "help": "Start from the defaults and choose everything yourself.",
        "values": {"name": "", "layout": "week_lanes", "weekdays": 127},
    },
}

WEEKDAY_PRESETS = {"every": 127, "work": 31, "weekend": 96}  # Monday is bit 0
