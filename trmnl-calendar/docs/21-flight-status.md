# 21. Flight status on events (design note)

Status: **built and verified 2026-11-15** (owner accepted FL-1 to FL-5). First live Check now: AC663 on
2026-11-16 read as YHZ → YUL, scheduled 11:25, gate 20, On time (R-15 retired). Decision record: ADR-0015.

## 1. What this is for

The owner's wife travels for work. Her flights are already events, drawn as they are today:

```
14:05 AC123        [M]
```

When an event holds a flight code, the app looks up that flight and adds **one word of status** to the
end of the same line, the way In / Out notes already do ("Work (out)"). Nothing else on the event changes:
the time is still the event's own start time, the badge stays where it is.

```
today, or no data yet     on time                  delayed
+--------------+          +--------------+         +--------------+
| 14:05 AC123 M|          | 14:05 AC123 M|         | 14:05 AC123 M|
|              |          | (On time)    |         | (Delayed)    |
+--------------+          +--------------+         +--------------+
```

A week cell is narrow, so the status usually wraps to the second line the cell already allows.

No alerts, no email, no SMS (owner, 2026-11-15). The screen and the web pages only.

## 2. Spotting a flight

- The **title or location** is searched for an airline code followed by a flight number:
  `\b([A-Z]{2}|[A-Z][0-9]|[0-9][A-Z])\s?([0-9]{1,4})\b`, upper case only. It matches `AC123`, `AC 123`,
  `WS456`, `B6 789`, `U2 1234`. Lower case ("ac123") is not matched, which keeps out most ordinary words.
  Two more rules, found by trying the pattern on sample titles: `AM` and `PM` are never airline codes
  (`PM 5`), and a code followed by a **space** and a year from 1970 to 2100 is skipped (`CA 2026 trip`).
  Written without the space (`CA2026`) it is still looked up.
- The code is looked up on the **day of the occurrence it is drawn on**. On a multi-day event (a trip),
  the **first** code belongs to the first day and the **last** code to the last day, so a trip titled
  `Toronto AC604 / AC611` gets the outbound status on day one and the return on the last day.
- A single-day event with two codes (a connection) shows the status of the first flight until it has
  landed, then the second.
- A match that is not a real flight costs **one** lookup: the "not found" answer is stored for that date
  and never asked again. Nothing is shown for it.

## 3. When the app looks a flight up

Flight data costs money per call, so lookups follow the flight, not the screen's 15-minute timer.

| Window | How often |
|--------|-----------|
| The day before (any time) | Once, to confirm the flight and learn its times in UTC |
| 4 hours before scheduled departure until it departs | Every 30 minutes |
| Departed until it arrives | Every 60 minutes, plus once at the expected arrival |
| Arrived, cancelled or diverted | Never again |
| 6 hours after scheduled arrival with no answer | Stop |

About 12 calls per flight. The scheduler loop (ADR-0007, one worker) asks the flight module for due
lookups before each render; each stored flight has a `next_due_at`. A monthly cap
(**Lookups a month** on the Settings page, default 100 calls) stops all lookups when reached; the last known status stays
on screen and the Screen page says the cap was hit.

The image is still a function of (data, settings, now) (NFR-7): the status comes from the database, not
from a call made while drawing. A changed status changes the picture, so the normal push-on-change sends
it. An unchanged status costs nothing.

## 4. The words

Status from the provider, mapped to what is shown. A delay is 15 minutes or more between the scheduled and
the revised departure (the usual airline threshold).

| Provider status | Shown |
|-----------------|-------|
| Expected, CheckIn, Boarding, GateClosed, Unknown with times; less than 15 minutes late | **On time** |
| Delayed, or any of the above 15 minutes or more late | **Delayed 2p** (with the new departure time when known) |
| Departed, EnRoute, Approaching | **Departed** |
| Arrived | **Landed** |
| Canceled, CanceledUncertain | **Cancelled** |
| Diverted | **Diverted** |
| No answer yet, not found, outside the window | nothing |

The status is written as a suffix in brackets, like notes: `12p AC123 (Delayed 2p)`. The event's own
time stays as entered; the new time is the departure airport's clock in the calendar's 12- or 24-hour
style, and a later new time is a change that is pushed like a new word. When the line would be
shortened to fit, the **title** is shortened and the status is kept, because the status is the point.
On a trip's bar, the status of the first day shown is used. The web pages (home list, month page, Screen
preview) and `GET /api/v1/occurrences` (field `flight`) show the same suffix. The event page has a
**Flight status** panel: for each code, the word, the route, scheduled and revised departure, gate when
known, the last failure if any, and when it was last checked; and a **Check now** button
(`POST /events/{id}/flights/check`) that asks once now, whatever the schedule, within the monthly cap.
When the feature is off the panel says which `.env` settings turn it on. Times there are the airport's
local times, which matches floating local time (ADR-0010).

**Filling a missing time** (owner, 2026-11-15, FL-6). When a lookup finds a flight and the event has
**no start time**, the event's start time is set to the planned departure (`sched_dep_local`, the
airport's clock), saved like an edit (its revision goes up, so it shows on the event page and can be
changed back). Only one-off events, only the event's first flight on its first day, and **never a time
someone typed**. A delay never moves it; the new time shows in the brackets. This is the one place the
feature writes to an event.

**Flights landing at home** (owner, 2026-11-20, FL-7). The Settings page has a **Home airport** (IATA
`YHZ` or ICAO `CYHZ`). A flight whose destination is that airport follows its **arrival** instead of its
departure:
- a missing start time is filled with the **planned arrival** (FL-6 still applies: a typed time stays);
- On time / Delayed is judged on the arrival (15 minutes or more late is Delayed, also while in the air),
  and a delay shows the new arrival time: `8:40p AC664 (Delayed 9:15p)`;
- Departed, Landed, Cancelled and Diverted are unchanged;
- the event goes on the day she **lands**. Lookups ask for flights departing *or* arriving that day
  (`dateLocalRole=Both`) and prefer the leg landing at home that day, so an overnight flight works.

Every other flight follows its departure, as before. The event page says "lands at home" and shows the
arrival times for such a flight.

## 5. Data

Migration 0009, one table. Derived data: not part of `GET /export`, and **left out of backups** (the
backup copy empties the flight tables), because the provider allows keeping its data for 7 days only and
backups are kept longer. Rows are deleted 7 days after `checked_at`; a status that is still needed is
simply fetched again.

`flight_status(code, dep_date, found, status, sched_dep, revised_dep, sched_arr, revised_arr,
sched_dep_local, revised_dep_local, dep_airport, arr_airport, dep_gate, checked_at, next_due_at,
last_error, sched_arr_local, revised_arr_local, dep_icao, arr_icao)`, primary key `(code, dep_date)`; times
in UTC, `*_local` is `HH:MM` on the airport's clock. Migration 0010 added the last four (2026-11-20).
`dep_date` is the event's day, which for a flight landing at home is the arrival day.
A failed lookup keeps the last good answer, records `last_error` and retries in 30 minutes. `code` is normalised without the space (`AC123`). A month counter for the cap
lives in `flight_calls(month, count)`.

## 6. The provider

**AeroDataBox** (through RapidAPI), endpoint `GET /flights/number/{code}/{dateLocal}`. It returns one or
more legs with scheduled and revised times, airports, gate and a status from the list in section 4. Its
free **Basic** plan (owner's screenshot, 2026-11-15): 400 units and 1,600 requests a month, both hard
limits; a tier-2 call such as flight status costs **2 units** (4 if flight plans are asked for, which they
are not). So 400 units buy **200 lookups**. The default cap of 100 lookups uses half and covers about
8 flights a month at about 12 lookups each; raise **Lookups a month** to at most 190 for about 15.

Plan terms that shape the design: **attribution is required**, so the event page, the home list and the
month page say "Flight data: AeroDataBox" whenever a flight status is shown (the e-paper image and the
Screen page preview do not, for space; to be checked against the terms). **Caching is limited to 7 days**, hence section 5.

Only `flight_source.py` talks to the service (as `holiday_source.py` does for holidays): standard library,
short timeout, size cap, one function that returns a small `FlightInfo` or "not found". A different
provider is a new version of that one file.

Settings follow ADR-0005. In the environment (template in `app/.env.example`): `FLIGHT_STATUS=off` by
default (it needs a key), `FLIGHT_API_KEY` (a secret: never logged, never in the repo), `FLIGHT_API_URL`.
On the **Settings** page (stored in `calendars.options_json`): **Home airport** and **Lookups a month**.
`FLIGHT_HOME_AIRPORT` and `FLIGHT_MONTHLY_LIMIT` in the environment only fill those two the first time the
app starts; once saved on the page, the page wins (moved 2026-11-20). Only the flight code and the date
are sent.

## 7. Decisions (owner, 2026-11-15 and 2026-11-20)

| # | Question | Decision |
|---|----------|----------|
| FL-7 | Flights coming home | A **Home airport** setting; flights landing there follow arrival (owner, 2026-11-20). |
| FL-8 | Where the flight preferences live | Home airport and lookups a month on the Settings page, per ADR-0005 (owner, 2026-11-20). |
| FL-6 | Set the start time from the flight? | Only when the event has none (owner, 2026-11-15); a typed time is never replaced. |
| FL-1 | Words only, or "Delayed 14:50" with the new time? | First words only; changed the same day (owner, 2026-11-15) to **the new time after Delayed**, keeping the event's own time. |
| FL-2 | Spot codes automatically, or only in one category such as "Travel"? | Automatically; a false match costs one call. |
| FL-3 | Status in brackets like notes, `(On time)`? | Yes, for consistency. |
| FL-4 | Provider | AeroDataBox; FlightAware AeroAPI if its data proves better (section 9 of ADR-0015). |
| FL-5 | Show "Departed" and "Landed" after take-off, or stop showing anything? | Show them; they tell the family she is in the air or down. |

A mock-up with the real renderer (AC663, Monday 11:25, badge M) showed the status wrapping to the
second line of a week cell with no shortening and no fit warnings.

## 8. Tests

Unit: the pattern (matches and near misses such as `Room B2`, `PM 5`, `CA 2026 trip`, `ac123`), first / last code on a trip,
connection order, status mapping and the 15-minute line, the lookup schedule against a fake clock, the
monthly cap, the title shortened before the status. Contract: recorded AeroDataBox replies (made-up
flights) for on time, delayed, cancelled, not found, and a garbled reply. Render: a golden image with
`(Delayed)` on a week cell. No test calls the live service.

## 9. Not in scope

Alerts of any kind, gate changes on the screen, arrival times on the screen, airport maps, tracking flights
that are not on the calendar.
