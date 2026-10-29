"""Event -> API JSON (docs/06 section 2). Pure."""

from __future__ import annotations

from trmnl_calendar.domain.models import Event
from trmnl_calendar.domain.repeat_text import repeat_text


def event_to_api(ev: Event, meta: dict | None = None, people: dict | None = None) -> dict:
    meta = meta or {}
    by = meta.get("updated_by_person_id")
    who = people.get(by) if people and by else None
    return {
        "id": ev.id,
        "uid": ev.uid,
        "rev": ev.rev,
        "title": ev.title,
        "category_id": ev.section_id,
        "start_date": ev.start_date.isoformat(),
        "end_date": ev.end_date.isoformat() if ev.end_date else None,
        "start_time": ev.start_time,
        "end_time": ev.end_time,
        "location": ev.location,
        "repeat": {
            "freq": ev.freq,
            "interval": ev.interval,
            "weekdays": list(ev.by_weekday),
            "month_mode": ev.month_mode,
            "month_nth": ev.month_nth,
            "until": ev.until.isoformat() if ev.until else None,
            "count": ev.count,
            "rotation": [list(w) for w in ev.rotation] if ev.rotation else None,
        },
        "repeat_text": repeat_text(ev),
        "people": {"ids": list(ev.people), "owner_id": ev.owner_id},
        "updated_at": meta.get("updated_at"),
        "updated_by": {"id": who.id, "label": who.label} if who else None,
        "deleted_at": meta.get("deleted_at"),
    }


def patch_from_api(body: dict) -> dict:
    """Map an API patch body onto repository field names (category_id -> section_id, repeat.*)."""
    patch = {k: v for k, v in body.items() if k not in ("repeat", "category_id", "people")}
    ppl = body.get("people")
    if isinstance(ppl, dict):
        if "ids" in ppl:
            patch["people_ids"] = ppl["ids"]
        if "owner_id" in ppl:
            patch["owner_id"] = ppl["owner_id"]
    if "category_id" in body:
        patch["section_id"] = body["category_id"]
    rep = body.get("repeat")
    if isinstance(rep, dict):
        mapping = {
            "freq": "freq",
            "interval": "interval",
            "weekdays": "by_weekday",
            "month_mode": "month_mode",
            "month_nth": "month_nth",
            "until": "until",
            "count": "count",
            "rotation": "rotation",
        }
        for src, dst in mapping.items():
            if src in rep:
                patch[dst] = rep[src]
    return patch
