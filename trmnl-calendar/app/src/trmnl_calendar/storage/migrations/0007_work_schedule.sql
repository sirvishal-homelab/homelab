-- Rotating-week schedules (e.g. "week A works Mon/Tue/Fri-Sun, week B works Wed/Thu") and a short
-- note on a single day of a repeating event (e.g. "in" / "out"). Both optional.
ALTER TABLE events ADD COLUMN rotation TEXT;  -- JSON: [["MO","TU"], ["WE","TH"]]; NULL = none

CREATE TABLE event_day_notes (
    event_id INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
    day      TEXT    NOT NULL,   -- the occurrence's start date, YYYY-MM-DD
    note     TEXT    NOT NULL CHECK (note IN ('in', 'out')),
    PRIMARY KEY (event_id, day)
) STRICT, WITHOUT ROWID;
