"""SQLite storage: events, sections, settings. Also validation and recurrence expansion."""
import json
import os
import sqlite3
from contextlib import closing
from datetime import date, datetime, timedelta

DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "data"))
DB_PATH = os.path.join(DATA_DIR, "events.db")

REPEATS = ("none", "daily", "weekly", "monthly", "yearly")
VIEWS = ("week", "twoweek", "month", "agenda", "split")
VIEW_LABELS = {
    "week": "Week",
    "twoweek": "Two weeks",
    "month": "Month",
    "agenda": "Agenda",
    "split": "Month and agenda",
}
DEFAULT_SECTIONS = ("Events", "Work", "Lunch", "Dinner")
MAX_SECTIONS = 8

SCHEMA = """
CREATE TABLE IF NOT EXISTS events (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    title        TEXT NOT NULL,
    date         TEXT NOT NULL,          -- first day, YYYY-MM-DD
    end_date     TEXT,                   -- last day for multi-day events
    start_time   TEXT,                   -- HH:MM, NULL = all day
    end_time     TEXT,                   -- HH:MM, optional
    location     TEXT NOT NULL DEFAULT '',
    repeat       TEXT NOT NULL DEFAULT 'none',
    repeat_until TEXT,                   -- last allowed start date, optional
    section_id   INTEGER                 -- which row of the day cell it appears in
);
CREATE INDEX IF NOT EXISTS idx_events_date ON events(date);

CREATE TABLE IF NOT EXISTS sections (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    name     TEXT NOT NULL,
    position INTEGER NOT NULL DEFAULT 0,
    weight   INTEGER NOT NULL DEFAULT 1,  -- relative height, 1-4
    visible  INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL                   -- JSON
);
"""


def _connect():
    os.makedirs(DATA_DIR, exist_ok=True)
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db


def init_db():
    """Create tables and migrate older databases (adds events.section_id)."""
    with closing(_connect()) as db:
        db.executescript(SCHEMA)
        cols = {r["name"] for r in db.execute("PRAGMA table_info(events)")}
        if "section_id" not in cols:
            db.execute("ALTER TABLE events ADD COLUMN section_id INTEGER")
        if db.execute("SELECT COUNT(*) FROM sections").fetchone()[0] == 0:
            for i, name in enumerate(DEFAULT_SECTIONS):
                db.execute("INSERT INTO sections (name, position) VALUES (?, ?)", (name, i))
        first = db.execute("SELECT id FROM sections ORDER BY position, id LIMIT 1").fetchone()[0]
        db.execute(
            "UPDATE events SET section_id = ? "
            "WHERE section_id IS NULL OR section_id NOT IN (SELECT id FROM sections)",
            (first,),
        )
        db.commit()


# ------------------------------------------------------------------ settings

def _env_choice(name, default, allowed):
    value = os.environ.get(name, default).lower()
    return value if value in allowed else default


# Environment variables only seed the defaults; once saved in the database, the database wins.
DEFAULT_SETTINGS = {
    "view": "month",
    "week_start": _env_choice("WEEK_START", "sun", ("sun", "mon")),
    "time_format": _env_choice("TIME_FORMAT", "12", ("12", "24")),
    "show_weekends": True,
    "show_section_labels": True,
    "hide_empty_sections": False,
    "show_updated": True,
}
SETTING_CHOICES = {
    "view": VIEWS,
    "week_start": ("sun", "mon"),
    "time_format": ("12", "24"),
}
BOOL_SETTINGS = ("show_weekends", "show_section_labels", "hide_empty_sections", "show_updated")


def get_settings():
    with closing(_connect()) as db:
        rows = db.execute("SELECT key, value FROM settings").fetchall()
    stored = {}
    for r in rows:
        if r["key"] in DEFAULT_SETTINGS:
            try:
                stored[r["key"]] = json.loads(r["value"])
            except ValueError:
                pass
    merged = {**DEFAULT_SETTINGS, **stored}
    # guard against bad stored values
    for key, allowed in SETTING_CHOICES.items():
        if merged[key] not in allowed:
            merged[key] = DEFAULT_SETTINGS[key]
    return merged


def validate_settings(payload):
    p = payload or {}
    clean, errors = {}, {}
    for key, allowed in SETTING_CHOICES.items():
        if key in p:
            if p[key] in allowed:
                clean[key] = p[key]
            else:
                errors[key] = "Choose one of the listed options."
    for key in BOOL_SETTINGS:
        if key in p:
            if isinstance(p[key], bool):
                clean[key] = p[key]
            else:
                errors[key] = "Must be on or off."
    return clean, errors


def update_settings(clean):
    with closing(_connect()) as db:
        for key, value in clean.items():
            db.execute(
                "INSERT INTO settings (key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (key, json.dumps(value)),
            )
        db.commit()
    return get_settings()


def get_internal(key, default=None):
    """Private values (e.g. the Terminus screen id). Not exposed through the settings API."""
    with closing(_connect()) as db:
        row = db.execute("SELECT value FROM settings WHERE key = ?", ("_" + key,)).fetchone()
    if not row:
        return default
    try:
        return json.loads(row["value"])
    except ValueError:
        return default


def set_internal(key, value):
    with closing(_connect()) as db:
        db.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            ("_" + key, json.dumps(value)),
        )
        db.commit()


# ------------------------------------------------------------------ sections

def _section(row):
    return {"id": row["id"], "name": row["name"], "position": row["position"],
            "weight": row["weight"], "visible": bool(row["visible"])}


def list_sections():
    with closing(_connect()) as db:
        rows = db.execute("SELECT * FROM sections ORDER BY position, id").fetchall()
    return [_section(r) for r in rows]


def _validate_section(fields, exclude_id=None):
    errors = {}
    name = str(fields.get("name") or "").strip()
    if not name:
        errors["name"] = "Enter a name."
    elif len(name) > 30:
        errors["name"] = "Keep the name under 30 characters."
    else:
        for s in list_sections():
            if s["id"] != exclude_id and s["name"].lower() == name.lower():
                errors["name"] = "You already have a section with that name."
    weight = fields.get("weight", 1)
    if isinstance(weight, bool) or not isinstance(weight, int) or not 1 <= weight <= 4:
        errors["weight"] = "Height must be between 1 and 4."
    visible = fields.get("visible", True)
    if not isinstance(visible, bool):
        errors["visible"] = "Must be on or off."
    return {"name": name, "weight": weight, "visible": visible}, errors


def create_section(payload):
    clean, errors = _validate_section(payload or {})
    if not errors and len(list_sections()) >= MAX_SECTIONS:
        errors["name"] = f"You can have up to {MAX_SECTIONS} sections."
    if errors:
        return None, errors
    with closing(_connect()) as db:
        pos = db.execute("SELECT COALESCE(MAX(position), -1) + 1 FROM sections").fetchone()[0]
        cur = db.execute(
            "INSERT INTO sections (name, position, weight, visible) VALUES (?, ?, ?, ?)",
            (clean["name"], pos, clean["weight"], int(clean["visible"])),
        )
        db.commit()
        new_id = cur.lastrowid
    return next(s for s in list_sections() if s["id"] == new_id), {}


def update_section(section_id, payload):
    """Partial update. Returns (section, errors); section is None if not found."""
    current = next((s for s in list_sections() if s["id"] == section_id), None)
    if not current:
        return None, {}
    merged = {k: (payload or {}).get(k, current[k]) for k in ("name", "weight", "visible")}
    clean, errors = _validate_section(merged, exclude_id=section_id)
    if errors:
        return current, errors
    with closing(_connect()) as db:
        db.execute("UPDATE sections SET name = ?, weight = ?, visible = ? WHERE id = ?",
                   (clean["name"], clean["weight"], int(clean["visible"]), section_id))
        db.commit()
    return next(s for s in list_sections() if s["id"] == section_id), {}


def reorder_sections(ids):
    existing = [s["id"] for s in list_sections()]
    if not isinstance(ids, list) or sorted(ids) != sorted(existing):
        return False
    with closing(_connect()) as db:
        for pos, sid in enumerate(ids):
            db.execute("UPDATE sections SET position = ? WHERE id = ?", (pos, sid))
        db.commit()
    return True


def delete_section(section_id):
    """Events in the deleted section move to the first remaining one. Returns (ok, message)."""
    sections = list_sections()
    if not any(s["id"] == section_id for s in sections):
        return False, "Section not found."
    remaining = [s for s in sections if s["id"] != section_id]
    if not remaining:
        return False, "Keep at least one section."
    with closing(_connect()) as db:
        db.execute("UPDATE events SET section_id = ? WHERE section_id = ?", (remaining[0]["id"], section_id))
        db.execute("DELETE FROM sections WHERE id = ?", (section_id,))
        db.commit()
    return True, ""


# ---------------------------------------------------------------- validation

def _parse_date(value):
    try:
        return date.fromisoformat(value)
    except (TypeError, ValueError):
        return None


def _parse_time(value):
    """Return 'HH:MM', None for blank, or False when invalid."""
    if not value:
        return None
    try:
        return datetime.strptime(value, "%H:%M").strftime("%H:%M")
    except ValueError:
        return False


def validate(payload):
    """Return (clean_dict, errors_dict) for an event."""
    p = payload or {}
    errors, clean = {}, {}

    title = str(p.get("title") or "").strip()
    if not title:
        errors["title"] = "Enter a title."
    elif len(title) > 120:
        errors["title"] = "Keep the title under 120 characters."
    clean["title"] = title

    start = _parse_date(p.get("date"))
    if not start:
        errors["date"] = "Choose a date."
    clean["date"] = start.isoformat() if start else None

    end = None
    if p.get("end_date"):
        end = _parse_date(p["end_date"])
        if not end:
            errors["end_date"] = "Choose a valid end date."
        elif start and end < start:
            errors["end_date"] = "The end date can't be before the start date."
            end = None
        elif start and end == start:
            end = None
    clean["end_date"] = end.isoformat() if end else None

    if p.get("all_day"):
        st = et = None
    else:
        st, et = _parse_time(p.get("start_time")), _parse_time(p.get("end_time"))
        if st is False:
            errors["start_time"] = "Use a time like 14:30."
            st = None
        if et is False:
            errors["end_time"] = "Use a time like 15:30."
            et = None
        if et and not st:
            errors["end_time"] = "Add a start time first."
            et = None
        elif st and et and not end and et <= st:
            errors["end_time"] = "The end time must be after the start time."
            et = None
    clean["start_time"], clean["end_time"] = st, et

    location = str(p.get("location") or "").strip()
    if len(location) > 120:
        errors["location"] = "Keep the location under 120 characters."
    clean["location"] = location

    repeat = p.get("repeat") or "none"
    if repeat not in REPEATS:
        errors["repeat"] = "Choose how often this repeats."
        repeat = "none"
    clean["repeat"] = repeat

    until = None
    if repeat != "none" and p.get("repeat_until"):
        until = _parse_date(p["repeat_until"])
        if not until:
            errors["repeat_until"] = "Choose a valid date."
        elif start and until < start:
            errors["repeat_until"] = "Repeat-until can't be before the first date."
            until = None
    clean["repeat_until"] = until.isoformat() if until else None

    sections = list_sections()
    ids = {s["id"] for s in sections}
    raw = p.get("section_id")
    if raw in (None, ""):
        clean["section_id"] = sections[0]["id"]
    else:
        try:
            sid = int(raw)
        except (TypeError, ValueError):
            sid = None
        if sid in ids:
            clean["section_id"] = sid
        else:
            errors["section_id"] = "Choose a section."
            clean["section_id"] = sections[0]["id"]

    return clean, errors


# ---------------------------------------------------------------------- CRUD

COLUMNS = ("title", "date", "end_date", "start_time", "end_time",
           "location", "repeat", "repeat_until", "section_id")


def list_events():
    with closing(_connect()) as db:
        rows = db.execute("SELECT * FROM events ORDER BY date, start_time, title").fetchall()
    return [dict(r) for r in rows]


def get_event(event_id):
    with closing(_connect()) as db:
        row = db.execute("SELECT * FROM events WHERE id = ?", (event_id,)).fetchone()
    return dict(row) if row else None


def create_event(clean):
    with closing(_connect()) as db:
        cur = db.execute(
            f"INSERT INTO events ({', '.join(COLUMNS)}) VALUES ({', '.join('?' * len(COLUMNS))})",
            [clean[c] for c in COLUMNS],
        )
        db.commit()
        return cur.lastrowid


def update_event(event_id, clean):
    with closing(_connect()) as db:
        cur = db.execute(
            f"UPDATE events SET {', '.join(c + ' = ?' for c in COLUMNS)} WHERE id = ?",
            [clean[c] for c in COLUMNS] + [event_id],
        )
        db.commit()
        return cur.rowcount > 0


def delete_event(event_id):
    with closing(_connect()) as db:
        cur = db.execute("DELETE FROM events WHERE id = ?", (event_id,))
        db.commit()
        return cur.rowcount > 0


# ---------------------------------------------------------------- recurrence

def _starts(first, repeat, lo, hi, until):
    """Yield start dates of an event that fall within [lo, hi]."""
    if until:
        hi = min(hi, until)
    if hi < lo:
        return
    if repeat == "none":
        if lo <= first <= hi:
            yield first
    elif repeat in ("daily", "weekly"):
        step = 1 if repeat == "daily" else 7
        d = first
        if lo > first:
            d = first + timedelta(days=-(-(lo - first).days // step) * step)
        while d <= hi:
            yield d
            d += timedelta(days=step)
    elif repeat == "monthly":
        y, m = first.year, first.month
        if lo > first:
            y, m = lo.year, lo.month
        while (y, m) <= (hi.year, hi.month):
            try:
                d = date(y, m, first.day)
                if lo <= d <= hi and d >= first:
                    yield d
            except ValueError:  # e.g. the 31st in a 30-day month
                pass
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    elif repeat == "yearly":
        for y in range(max(first.year, lo.year), hi.year + 1):
            try:
                d = date(y, first.month, first.day)
                if lo <= d <= hi and d >= first:
                    yield d
            except ValueError:  # Feb 29 in a non-leap year
                pass


def events_by_day(win_start, win_end):
    """Return {date: [entry, ...]} for every day in the window.

    Multi-day events appear on each day they cover; `is_start` is True only on
    the first day so times are shown once. All-day entries sort first.
    """
    result = {}
    for ev in list_events():
        first = date.fromisoformat(ev["date"])
        last = date.fromisoformat(ev["end_date"]) if ev["end_date"] else first
        span = (last - first).days
        until = date.fromisoformat(ev["repeat_until"]) if ev["repeat_until"] else None
        lo = win_start - timedelta(days=span)
        for s in _starts(first, ev["repeat"], lo, win_end, until):
            for i in range(span + 1):
                day = s + timedelta(days=i)
                if win_start <= day <= win_end:
                    result.setdefault(day, []).append({
                        "id": ev["id"],
                        "title": ev["title"],
                        "location": ev["location"],
                        "section_id": ev["section_id"],
                        "start_time": ev["start_time"] if i == 0 else None,
                        "end_time": ev["end_time"] if i == 0 else None,
                        "is_start": i == 0,
                        "all_day": ev["start_time"] is None or i > 0,
                        "day_number": i + 1,
                        "day_count": span + 1,
                    })
    for entries in result.values():
        entries.sort(key=lambda e: (e["start_time"] or "", e["title"].lower()))
    return result
