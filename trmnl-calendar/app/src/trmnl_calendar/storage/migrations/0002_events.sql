-- Slice 1: calendar, categories (sections), events, skipped dates, idempotency keys.
-- People and event_people arrive in slice 2; views and displays in slices 3 and 4.
CREATE TABLE calendars (
    id           INTEGER PRIMARY KEY,
    name         TEXT NOT NULL,
    timezone     TEXT NOT NULL,
    options_json TEXT NOT NULL DEFAULT '{}'
) STRICT;
INSERT INTO calendars (id, name, timezone) VALUES (1, 'Calendar', 'America/Halifax');

CREATE TABLE sections (
    id          INTEGER PRIMARY KEY,
    calendar_id INTEGER NOT NULL REFERENCES calendars(id),
    name        TEXT NOT NULL COLLATE NOCASE CHECK (length(name) BETWEEN 1 AND 30),
    position    INTEGER NOT NULL,
    weight      INTEGER NOT NULL DEFAULT 1 CHECK (weight BETWEEN 1 AND 4),
    visible     INTEGER NOT NULL DEFAULT 1 CHECK (visible IN (0, 1)),
    placement   TEXT NOT NULL DEFAULT 'lane' CHECK (placement IN ('lane', 'banner')),
    UNIQUE (calendar_id, name)
) STRICT;
INSERT INTO sections (calendar_id, name, position) VALUES
    (1, 'Events', 1), (1, 'Work', 2), (1, 'Lunch', 3), (1, 'Dinner', 4);

CREATE TABLE events (
    id                   INTEGER PRIMARY KEY,
    uid                  TEXT NOT NULL UNIQUE,
    calendar_id          INTEGER NOT NULL REFERENCES calendars(id),
    section_id           INTEGER NOT NULL REFERENCES sections(id),
    title                TEXT NOT NULL CHECK (length(title) BETWEEN 1 AND 120),
    location             TEXT NOT NULL DEFAULT '' CHECK (length(location) <= 120),
    start_date           TEXT NOT NULL,
    end_date             TEXT,
    start_time           TEXT,
    end_time             TEXT,
    freq                 TEXT NOT NULL DEFAULT 'none'
                         CHECK (freq IN ('none', 'daily', 'weekly', 'monthly', 'yearly')),
    interval             INTEGER NOT NULL DEFAULT 1 CHECK (interval BETWEEN 1 AND 99),
    by_weekday           TEXT NOT NULL DEFAULT '',
    month_mode           TEXT CHECK (month_mode IN ('date', 'weekday')),
    month_nth            INTEGER CHECK (month_nth IN (1, 2, 3, 4, -1)),
    until                TEXT,
    rev                  INTEGER NOT NULL DEFAULT 1,
    updated_by_person_id INTEGER,
    created_at           TEXT NOT NULL,
    updated_at           TEXT NOT NULL,
    deleted_at           TEXT
) STRICT;
CREATE INDEX events_live ON events (calendar_id, start_date) WHERE deleted_at IS NULL;

CREATE TABLE event_exceptions (
    event_id        INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
    occurrence_date TEXT NOT NULL,
    kind            TEXT NOT NULL DEFAULT 'skip' CHECK (kind IN ('skip')),
    PRIMARY KEY (event_id, occurrence_date)
) STRICT;

CREATE TABLE idempotency_keys (
    key             TEXT NOT NULL,
    method_and_path TEXT NOT NULL,
    status          INTEGER NOT NULL,
    response_json   TEXT NOT NULL,
    created_at      TEXT NOT NULL,
    PRIMARY KEY (key, method_and_path)
) STRICT;
