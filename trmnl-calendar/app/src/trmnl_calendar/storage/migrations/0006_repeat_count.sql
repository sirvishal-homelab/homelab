-- Repeat "N times" as an alternative to "until a date". The first occurrence counts as one.
ALTER TABLE events ADD COLUMN repeat_count INTEGER
    CHECK (repeat_count IS NULL OR repeat_count BETWEEN 2 AND 999);
