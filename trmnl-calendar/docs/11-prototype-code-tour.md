# 11 - Prototype code tour

The prototype lives in `reference/prototype/`. This page maps it to the design and shows the
parts most relevant to design review. Excerpts are copied from the files and may be
**ahead of or behind** the rest of the prototype; see `reference/prototype/STATUS.md`.

## File map

| File | Responsibility | Design doc |
|------|----------------|------------|
| `store.py` | SQLite access, migration, events/sections/settings CRUD, validation, recurrence expansion | `03` |
| `render.py` | Canvas, fonts, section lanes, week/two-week/month/agenda/split layouts, 1-bit PNG | `04` |
| `publisher.py` | Terminus client and background push worker | `05` |
| `app.py` | Flask routes (**v1: events, image, upcoming only**) | `06` |
| `templates/index.html`, `static/*` | Editor UI (**v1**) | `07` |
| `seed_demo.py` | Sample data across sections | n/a |
| `Dockerfile`, `docker-compose.yml` | Container setup (**needs rework**) | `02` |

Prototype dependencies: Flask, Pillow, gunicorn, tzdata. Standard library only for HTTP
(`urllib`) and SQLite (`sqlite3`).

## Database schema (from `store.py`)

```sql
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
```

## Default settings (from `store.py`)

```python
DEFAULT_SETTINGS = {
    "view": "month",
    "week_start": _env_choice("WEEK_START", "sun", ("sun", "mon")),
    "time_format": _env_choice("TIME_FORMAT", "12", ("12", "24")),
    "show_weekends": True,
    "show_section_labels": True,
    "hide_empty_sections": False,
    "show_updated": True,
}
```

## Section lane selection (from `render.py`)

```python
def pick_sections(sections, by_day, dates, cfg):
    """Sections to draw: only visible ones, and optionally only those with events in range."""
    ids = {s["id"] for s in sections}
    default_id = sections[0]["id"]
    chosen = [s for s in sections if s["visible"]]
    if cfg["hide_empty_sections"]:
        used = set()
        for d in dates:
            for sid in bucket_by_section(by_day.get(d, []), ids, default_id):
                used.add(sid)
        chosen = [s for s in chosen if s["id"] in used] or chosen[:1]
    return chosen or sections[:1]


def lane_heights(chosen, total):
    weights = [s["weight"] for s in chosen]
    unit = total / float(sum(weights))
    heights, y = [], 0.0
    for wgt in weights:
        heights.append((y, wgt * unit))
        y += wgt * unit
    return heights
```

## Entry point of the renderer (from `render.py`)

```python
def render_png(view, w, h, now, cfg=None, sections=None):
    cfg = cfg or store.get_settings()
    sections = sections or store.list_sections()
    view = view if view in VIEWS else cfg["view"]
    today = now.date()
    c = Canvas(w, h)
    m = c.px(22)

    first_of_month = today.replace(day=1)
    win_start = min(first_of_month - timedelta(days=14), today - timedelta(days=7))
    win_end = max(first_of_month + timedelta(days=45), today + timedelta(days=AGENDA_DAYS))
    by_day = store.events_by_day(win_start, win_end)

    stamp = None
    if cfg["show_updated"]:
        t = now.strftime("%H:%M") if cfg["time_format"] == "24" else now.strftime("%-I:%M %p")
        stamp = f"Updated {t}"
    foot = c.px(14) if stamp else 0

    if view in ("week", "twoweek", "month"):
        draw_grid(c, m, m, w - 2 * m, h - 2 * m - foot, view, today, by_day, sections, cfg)
    elif view == "agenda":
        draw_agenda(c, m, m, w - 2 * m, h - 2 * m - c.px(14), today, by_day, cfg, sections,
                    heading_right=f"{pycal.day_name[today.weekday()]} {today.day} {pycal.month_name[today.month]}")
    else:
        left_w = int(w * 0.43)
        draw_month_compact(c, m, m, left_w - m, h - 2 * m, today, by_day, cfg)
        div_x = left_w + c.px(10)
        c.vline(div_x, m, h - m)
        ax = div_x + c.px(22)
        draw_agenda(c, ax, m, w - ax - m, h - 2 * m - c.px(14), today, by_day, cfg, sections)
    if stamp:
        c.text((w - m, h - m + c.px(4)), stamp, 11, anchor="rs")
    return c.png()
```

## Terminus push (from `publisher.py`)

Untested. Shown so reviewers can judge the intended flow, not because it is proven.

```python
def push(self, reason="manual"):
    """Send the current image to Terminus. Returns the status dict; never raises."""
    if not self.configured():
        self._set("error", "Terminus is not configured. Missing: " + ", ".join(self.missing()), reason)
        return self.status()
    with self._lock:
        try:
            token = self._login()
            screen_id = store.get_internal("terminus_screen_id")
            if screen_id:
                status, resp = self._authed("PATCH", f"/api/screens/{screen_id}", self._screen_payload(False), token)
                if status == 404:
                    screen_id = None          # screen was deleted in Terminus; recreate it
            if not screen_id:
                if not self.model_id:
                    raise PushError("TERMINUS_MODEL_ID is required to create the screen the first time.")
                body = self._screen_payload(True)
                if self.playlist_id:
                    body["screen"]["playlist_id"] = self.playlist_id
                status, resp = self._authed("POST", "/api/screens", body, token)
            if status == 401:
                raise PushError("Terminus rejected the login token (HTTP 401).")
            if status >= 300:
                detail = _problem(resp) or f"HTTP {status}"
                raise PushError(f"Terminus refused the screen update: {detail}")
            new_id = ((resp or {}).get("data") or {}).get("id")
            if new_id:
                store.set_internal("terminus_screen_id", new_id)
            self._set("ok", "Sent the latest calendar image to Terminus.", reason)
        except PushError as err:
            self._set("error", str(err), reason)
        except Exception as err:  # keep the worker alive whatever happens
            self._set("error", f"Unexpected error: {err}", reason)
    return self.status()
```

## Things the next agent may want to reuse as **ideas**, not code

- The lane model (weights, visibility, hide-empty) and the rule "events in an unknown section
  fall back to the first section".
- Rendering lessons in `04` (bars instead of inverted small text).
- The idea of env vars seeding defaults while the database holds the saved values.
- The 3-second debounce plus periodic push in the publisher.

## Things the next agent should probably **not** carry over

- Everything being in four large modules with module-level globals.
- `PUT` as full replace.
- Inline `ALTER TABLE` migration.
- Two gunicorn workers.
