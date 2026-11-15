-- Baseline schema (docs/05 section 7). Tuning lives in `settings`; secrets never do.

CREATE TABLE settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL  -- JSON
);

INSERT INTO settings (key, value) VALUES
    ('device_url', '""'),
    ('warn_pct', '75'),
    ('crit_pct', '90'),
    ('poll_s', '60'),
    ('alert_flip_s', '10'),
    ('unreachable_grace_min', '5'),
    ('container_missing_polls', '2');

CREATE TABLE systems (
    beszel_id     TEXT PRIMARY KEY,
    beszel_name   TEXT NOT NULL,
    display_name  TEXT,
    focus_order   INTEGER CHECK (focus_order IS NULL OR focus_order BETWEEN 1 AND 4),
    ignore_alerts INTEGER NOT NULL DEFAULT 0,
    first_seen    TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now')),
    last_seen     TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now'))
);
CREATE UNIQUE INDEX systems_focus_order ON systems (focus_order) WHERE focus_order IS NOT NULL;

CREATE TABLE known_containers (
    system_id       TEXT NOT NULL,
    name            TEXT NOT NULL,
    allowed_to_stop INTEGER NOT NULL DEFAULT 0,
    last_running    TEXT,
    missing_polls   INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (system_id, name)
);

CREATE TABLE push_state (
    id            INTEGER PRIMARY KEY CHECK (id = 1),
    last_hash     TEXT,
    last_push_at  TEXT,
    last_result   TEXT,
    last_error    TEXT
);
INSERT INTO push_state (id) VALUES (1);
