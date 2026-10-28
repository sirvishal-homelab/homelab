-- Slice 4: what the app last sent to Terminus, so the status panel survives restarts (ADR-0007).
CREATE TABLE push_state (
    display_id      INTEGER PRIMARY KEY REFERENCES displays(id),
    screen_id       INTEGER,
    last_hash       TEXT,
    last_status     TEXT NOT NULL DEFAULT 'never'
                    CHECK (last_status IN ('never', 'ok', 'unchanged', 'error', 'in_progress')),
    last_error      TEXT,
    last_attempt_at TEXT,
    last_pushed_at  TEXT,
    failures        INTEGER NOT NULL DEFAULT 0
) STRICT;

INSERT INTO push_state (display_id) VALUES (1);
