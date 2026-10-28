"""Displays and saved views: validation, CRUD, order, active view (docs/03, docs/06 section 3)."""

from __future__ import annotations

import json
import sqlite3

from trmnl_calendar.domain.models import WEEKDAY_CODES, Display, View
from trmnl_calendar.domain.view_templates import TEMPLATES
from trmnl_calendar.errors import InUse, NotFound, ValidationFailed
from trmnl_calendar.render.layouts import registry
from trmnl_calendar.storage.repositories import people as people_repo
from trmnl_calendar.storage.repositories.calendar import list_sections

DISPLAY_ID = 1  # v1 has one display (A-1)
MAX_VIEWS = 12


def get_display(conn: sqlite3.Connection) -> Display:
    r = conn.execute("SELECT * FROM displays WHERE id = ?", (DISPLAY_ID,)).fetchone()
    return Display(
        r["id"],
        r["name"],
        r["model_name"],
        r["width"],
        r["height"],
        r["bit_depth"],
        r["active_view_id"],
        bool(r["enabled"]),
    )


def _row_to_view(r: sqlite3.Row) -> View:
    return View(
        id=r["id"],
        name=r["name"],
        template=r["template"],
        layout=r["layout"],
        weekdays=r["weekdays"],
        categories_mode=r["categories_mode"],
        category_ids=tuple(json.loads(r["categories_json"])),
        people_mode=r["people_mode"],
        person_ids=tuple(json.loads(r["people_json"])),
        options=json.loads(r["options_json"]),
        position=r["position"],
    )


def list_views(conn: sqlite3.Connection) -> list[View]:
    rows = conn.execute(
        "SELECT * FROM views WHERE display_id = ? ORDER BY position, id", (DISPLAY_ID,)
    ).fetchall()
    return [_row_to_view(r) for r in rows]


def get_view(conn: sqlite3.Connection, view_id: int) -> View:
    r = conn.execute(
        "SELECT * FROM views WHERE id = ? AND display_id = ?", (view_id, DISPLAY_ID)
    ).fetchone()
    if r is None:
        raise NotFound("That view no longer exists.")
    return _row_to_view(r)


def active_view(conn: sqlite3.Connection) -> View:
    """The view on the screen; if none is set (or it vanished) the first view."""
    display = get_display(conn)
    views = list_views(conn)
    for v in views:
        if v.id == display.active_view_id:
            return v
    if not views:
        raise NotFound("There are no views.")
    return views[0]


def weekdays_to_codes(mask: int) -> list[str]:
    return [c for i, c in enumerate(WEEKDAY_CODES) if mask >> i & 1]


def codes_to_mask(codes: object) -> tuple[int, str]:
    if not isinstance(codes, list | tuple) or not all(isinstance(c, str) for c in codes):
        return 0, "Choose the days to show"
    bad = [c for c in codes if c not in WEEKDAY_CODES]
    if bad:
        return 0, "Unknown day: " + ", ".join(bad)
    mask = sum(1 << WEEKDAY_CODES.index(c) for c in set(codes))
    return mask, ("" if mask else "Choose at least one day")


def view_to_dict(v: View) -> dict:
    """The editable shape, shared by the API, forms and merge-on-update."""
    return {
        "name": v.name,
        "template": v.template,
        "layout": v.layout,
        "weekdays": weekdays_to_codes(v.weekdays),
        "categories": {"mode": v.categories_mode, "ids": list(v.category_ids)},
        "people": {"mode": v.people_mode, "ids": list(v.person_ids)},
        "options": dict(v.options),
    }


def _id_filter(conn, raw: object, label: str, valid: set[int]) -> tuple[str, tuple, str]:
    if raw is None:
        raw = {}
    if not isinstance(raw, dict) or set(raw) - {"mode", "ids"}:
        return "all", (), f"Send {label} as mode and ids"
    mode = raw.get("mode", "all")
    ids = raw.get("ids", [])
    if mode not in ("all", "only"):
        return "all", (), "Mode must be all or only"
    if not isinstance(ids, list) or not all(
        isinstance(i, int) and not isinstance(i, bool) for i in ids
    ):
        return mode, (), "ids must be a list of numbers"
    if set(ids) - valid:
        return mode, (), f"One of those {label} no longer exists"
    if mode == "only" and not ids:
        return mode, (), f"Choose at least one, or show all {label}"
    return mode, tuple(dict.fromkeys(ids)) if mode == "only" else (), ""


def build_view(
    conn: sqlite3.Connection, data: dict, *, view_id: int | None = None, position: int = 0
) -> View:
    """Validate a full editable dict into a View (not saved). Raises ValidationFailed."""
    err: dict[str, str] = {}
    unknown = set(data) - {
        "name",
        "template",
        "layout",
        "weekdays",
        "categories",
        "people",
        "options",
    }
    for k in sorted(unknown):
        err[k] = "Unknown field"

    name = data.get("name", "")
    name = name.strip() if isinstance(name, str) else ""
    if not 1 <= len(name) <= 40:
        err["name"] = "Enter a name up to 40 characters"
    else:
        clash = conn.execute(
            "SELECT id FROM views WHERE display_id = ? AND name = ? COLLATE NOCASE AND id IS NOT ?",
            (DISPLAY_ID, name, view_id),
        ).fetchone()
        if clash:
            err["name"] = "You already have a view with that name"

    layout = data.get("layout")
    if layout not in registry.LAYOUTS:
        err["layout"] = "Choose a layout"
        layout = "week_lanes"

    mask, msg = codes_to_mask(data.get("weekdays"))
    if msg:
        err["weekdays"] = msg

    c_mode, c_ids, msg = _id_filter(
        conn, data.get("categories"), "categories", {s.id for s in list_sections(conn)}
    )
    if msg:
        err["categories"] = msg
    p_mode, p_ids, msg = _id_filter(
        conn, data.get("people"), "people", {p.id for p in people_repo.list_people(conn)}
    )
    if msg:
        err["people"] = msg

    raw_opts = data.get("options") or {}
    options = registry.default_options(layout)
    if not isinstance(raw_opts, dict):
        err["options"] = "Send options as an object"
    else:
        for k, val in raw_opts.items():
            if k not in registry.option_keys(layout):
                continue  # an option of another layout (the layout was changed): dropped
            if not isinstance(val, bool):
                err["options"] = f"{k} must be true or false"
            else:
                options[k] = val

    template = data.get("template", "blank")
    if template not in TEMPLATES:
        template = "blank"
    if err:
        raise ValidationFailed(err)
    return View(
        id=view_id,
        name=name,
        template=template,
        layout=layout,
        weekdays=mask,
        categories_mode=c_mode,
        category_ids=c_ids,
        people_mode=p_mode,
        person_ids=p_ids,
        options=options,
        position=position,
    )


def template_values(template: str) -> dict:
    if template not in TEMPLATES:
        raise ValidationFailed({"template": "Unknown template"})
    t = TEMPLATES[template]["values"]
    return {
        "name": t["name"],
        "template": template,
        "layout": t["layout"],
        "weekdays": weekdays_to_codes(t["weekdays"]),
        "categories": {"mode": "all", "ids": []},
        "people": {"mode": "all", "ids": []},
        "options": {},
    }


def _write(conn: sqlite3.Connection, v: View, view_id: int | None) -> int:
    cols = (
        v.name,
        v.template,
        v.layout,
        v.weekdays,
        v.categories_mode,
        json.dumps(list(v.category_ids)),
        v.people_mode,
        json.dumps(list(v.person_ids)),
        json.dumps(v.options, sort_keys=True),
    )
    if view_id is None:
        cur = conn.execute(
            "INSERT INTO views (display_id, name, template, layout, weekdays, categories_mode, "
            "categories_json, people_mode, people_json, options_json, position) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (DISPLAY_ID, *cols, v.position),
        )
        view_id = cur.lastrowid
    else:
        conn.execute(
            "UPDATE views SET name=?, template=?, layout=?, weekdays=?, categories_mode=?, "
            "categories_json=?, people_mode=?, people_json=?, options_json=? WHERE id=?",
            (*cols, view_id),
        )
    conn.commit()
    return view_id


def create_view(
    conn: sqlite3.Connection, template: str = "blank", overrides: dict | None = None
) -> View:
    data = {**template_values(template), **(overrides or {})}
    pos = conn.execute(
        "SELECT COUNT(*), COALESCE(MAX(position), 0) FROM views WHERE display_id = ?", (DISPLAY_ID,)
    ).fetchone()
    if pos[0] >= MAX_VIEWS:
        raise ValidationFailed({"name": f"You can have up to {MAX_VIEWS} views"})
    v = build_view(conn, data, position=pos[1] + 1)
    return get_view(conn, _write(conn, v, None))


def update_view(conn: sqlite3.Connection, view_id: int, patch: dict) -> View:
    cur = get_view(conn, view_id)
    data = {**view_to_dict(cur), **patch}
    if "layout" in patch and patch["layout"] != cur.layout and "options" not in patch:
        data["options"] = {}  # new layout, its own defaults
    v = build_view(conn, data, view_id=view_id, position=cur.position)
    _write(conn, v, view_id)
    return get_view(conn, view_id)


def delete_view(conn: sqlite3.Connection, view_id: int) -> None:
    get_view(conn, view_id)
    if active_view(conn).id == view_id:
        raise InUse("That view is showing on the screen. Show another view first.")
    conn.execute("DELETE FROM views WHERE id = ?", (view_id,))
    conn.commit()


def reorder_views(conn: sqlite3.Connection, ids: list[int]) -> None:
    have = [v.id for v in list_views(conn)]
    if sorted(ids) != sorted(have):
        raise ValidationFailed({"ids": "Send every view id exactly once"})
    for pos, vid in enumerate(ids, start=1):
        conn.execute("UPDATE views SET position = ? WHERE id = ?", (pos, vid))
    conn.commit()


def activate_view(conn: sqlite3.Connection, view_id: int) -> View:
    v = get_view(conn, view_id)
    conn.execute("UPDATE displays SET active_view_id = ? WHERE id = ?", (view_id, DISPLAY_ID))
    conn.commit()
    return v
