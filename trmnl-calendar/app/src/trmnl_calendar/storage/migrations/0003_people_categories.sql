-- Slice 2: people (with a one- or two-letter label drawn on the screen) and who an event is for.
-- No people are seeded: the owner adds the household on the People page.
CREATE TABLE people (
    id          INTEGER PRIMARY KEY,
    calendar_id INTEGER NOT NULL REFERENCES calendars(id),
    name        TEXT NOT NULL CHECK (length(name) BETWEEN 1 AND 30),
    label       TEXT NOT NULL COLLATE NOCASE CHECK (length(label) BETWEEN 1 AND 2),
    position    INTEGER NOT NULL,
    active      INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
    UNIQUE (calendar_id, label)
) STRICT;

-- Zero rows for an event means everyone. At most one owner per event.
CREATE TABLE event_people (
    event_id  INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
    person_id INTEGER NOT NULL REFERENCES people(id),
    role      TEXT NOT NULL DEFAULT 'with' CHECK (role IN ('owner', 'with')),
    PRIMARY KEY (event_id, person_id)
) STRICT;
CREATE UNIQUE INDEX event_people_one_owner ON event_people (event_id) WHERE role = 'owner';
CREATE INDEX event_people_person ON event_people (person_id);
