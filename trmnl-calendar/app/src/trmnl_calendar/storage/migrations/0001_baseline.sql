-- Baseline: only bookkeeping. Domain tables arrive with the slice that first needs them.
CREATE TABLE app_meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
INSERT INTO app_meta (key, value) VALUES ('created_by', 'slice-0');
