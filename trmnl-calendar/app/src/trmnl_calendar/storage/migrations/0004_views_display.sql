-- Slice 3: the display (render target) and saved views. Two starter views are seeded so the screen
-- looks exactly as before until the owner builds another.
CREATE TABLE displays (
    id             INTEGER PRIMARY KEY,
    calendar_id    INTEGER NOT NULL REFERENCES calendars(id),
    name           TEXT NOT NULL,
    model_name     TEXT NOT NULL,
    width          INTEGER NOT NULL,
    height         INTEGER NOT NULL,
    bit_depth      INTEGER NOT NULL CHECK (bit_depth IN (1, 2, 4)),
    active_view_id INTEGER,
    enabled        INTEGER NOT NULL DEFAULT 1 CHECK (enabled IN (0, 1))
) STRICT;

CREATE TABLE views (
    id              INTEGER PRIMARY KEY,
    display_id      INTEGER NOT NULL REFERENCES displays(id),
    name            TEXT NOT NULL CHECK (length(name) BETWEEN 1 AND 40),
    template        TEXT NOT NULL DEFAULT 'blank',
    layout          TEXT NOT NULL,
    weekdays        INTEGER NOT NULL DEFAULT 127 CHECK (weekdays BETWEEN 1 AND 127),
    categories_mode TEXT NOT NULL DEFAULT 'all' CHECK (categories_mode IN ('all', 'only')),
    categories_json TEXT NOT NULL DEFAULT '[]',
    people_mode     TEXT NOT NULL DEFAULT 'all' CHECK (people_mode IN ('all', 'only')),
    people_json     TEXT NOT NULL DEFAULT '[]',
    options_json    TEXT NOT NULL DEFAULT '{}',
    position        INTEGER NOT NULL
) STRICT;

INSERT INTO displays (id, calendar_id, name, model_name, width, height, bit_depth)
VALUES (1, 1, 'TRMNL', 'og_plus', 800, 480, 2);

INSERT INTO views (id, display_id, name, template, layout, position) VALUES
    (1, 1, 'Week - lanes', 'week_lanes', 'week_lanes', 1),
    (2, 1, 'Two weeks - list', 'twoweek_list', 'twoweek_list', 2);

UPDATE displays SET active_view_id = 1 WHERE id = 1;
