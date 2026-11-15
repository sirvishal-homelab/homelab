-- Flight status for flight codes found in events (ADR-0015, docs/21). Derived data: the provider's
-- terms allow keeping it 7 days, so rows go 7 days after `checked_at`, backups leave this table empty,
-- and GET /export does not include it.
CREATE TABLE flight_status (
    code              TEXT NOT NULL,      -- airline code and number without a space, "AC663"
    dep_date          TEXT NOT NULL,      -- the event's day, YYYY-MM-DD (the airport's local date)
    found             INTEGER NOT NULL,   -- 0: the provider has no such flight that day
    status            TEXT NOT NULL DEFAULT '',  -- the provider's word: Expected, Delayed, Arrived...
    sched_dep         TEXT,               -- UTC, YYYY-MM-DDTHH:MM:SSZ
    revised_dep       TEXT,
    sched_arr         TEXT,
    revised_arr       TEXT,
    sched_dep_local   TEXT NOT NULL DEFAULT '',  -- HH:MM on the departure airport's clock
    revised_dep_local TEXT NOT NULL DEFAULT '',
    dep_airport       TEXT NOT NULL DEFAULT '',
    arr_airport       TEXT NOT NULL DEFAULT '',
    dep_gate          TEXT NOT NULL DEFAULT '',
    checked_at        TEXT NOT NULL,      -- UTC, last attempt
    next_due_at       TEXT,               -- UTC; NULL means never ask again
    last_error        TEXT NOT NULL DEFAULT '',  -- the last failure, for the event page
    PRIMARY KEY (code, dep_date)
) STRICT, WITHOUT ROWID;

-- Calls made per UTC month, for FLIGHT_MONTHLY_LIMIT. Not provider data, so backups keep it.
CREATE TABLE flight_calls (
    month TEXT PRIMARY KEY,   -- YYYY-MM
    count INTEGER NOT NULL
) STRICT, WITHOUT ROWID;
