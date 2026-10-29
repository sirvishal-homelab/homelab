-- Battery and Wi-Fi readings of the screen, taken from Terminus's device record, so Settings can show
-- a trend (Terminus itself only shows the latest value). Rows older than 90 days are removed.
CREATE TABLE device_samples (
    at        TEXT PRIMARY KEY,   -- when this app took the reading, UTC, YYYY-MM-DDTHH:MM:SSZ
    voltage   REAL NOT NULL,
    wifi      INTEGER,            -- signal strength in dBm, if Terminus reports it
    seen_at   TEXT                -- the device's own last check-in time, UTC, if reported
) STRICT, WITHOUT ROWID;
