# ADR-0015: Flight status on events
Status: **Proposed 2026-11-15** (owner idea; not built)
Date: 2026-11-15

## Context
The owner's wife travels for work and her flights are events (`14:05 AC123` with her badge). The family
wants to see at a glance whether the flight is on time or delayed, without opening an app. The owner
wants the event to look as it does now, with the status added. Alerts are not wanted.

## Decision
- An event whose title or location holds an upper-case airline code and flight number (`AC123`,
  `B6 789`) gets a status suffix on the occurrence for that flight's date: **On time**, **Delayed**,
  **Departed**, **Landed**, **Cancelled** or **Diverted** (`21`, section 4). Nothing else changes.
- Status is fetched from **AeroDataBox** by `flight_source.py` only, on a schedule that follows each
  flight (once the day before, every 30 minutes from 4 hours before departure, hourly in the air, never
  after arrival), with a monthly call cap. Results are stored in `flight_status` (migration 0009), so
  rendering stays deterministic and push-on-change does the rest.
- Off unless `FLIGHT_STATUS=on` and `FLIGHT_API_KEY` are set in the environment. Only the code and date
  leave the network.

## Consequences
Good: no new screen, no new form field; one small module and one table; a change of status reaches the
panel through the existing push; nothing changes for people who do not turn it on.
Bad: a paid outside service with a small free allowance, so the cap can run out in a heavy travel month;
the data is only as good as the provider's; a false match costs one call; this is designed without a live
call, so the first real flight is the verification (`08`, R-15).

## Alternatives considered
- **FlightAware AeroAPI**: the best data; the Personal tier gives a few dollars of free queries a month,
  then charges per query. The fallback if AeroDataBox proves unreliable.
- **AviationStack**: free plan is about 100 requests a month and plain HTTP only, which would send the key
  unencrypted. Rejected.
- **A "Flight" field on the event**: exact, no false matches, but a form change and one more thing to fill
  in. The pattern already matches how she writes flights.
- **Looking up on every render**: simplest, but about 96 calls a day per flight and it would break
  NFR-7. Rejected.
