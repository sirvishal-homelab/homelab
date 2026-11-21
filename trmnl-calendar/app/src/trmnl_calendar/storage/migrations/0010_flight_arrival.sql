-- Arrival on the airport's own clock and ICAO codes, so a flight landing at the calendar's home
-- airport can follow its arrival (docs/21, ADR-0015 addendum 2026-11-20).
ALTER TABLE flight_status ADD COLUMN sched_arr_local TEXT NOT NULL DEFAULT '';
ALTER TABLE flight_status ADD COLUMN revised_arr_local TEXT NOT NULL DEFAULT '';
ALTER TABLE flight_status ADD COLUMN dep_icao TEXT NOT NULL DEFAULT '';
ALTER TABLE flight_status ADD COLUMN arr_icao TEXT NOT NULL DEFAULT '';
