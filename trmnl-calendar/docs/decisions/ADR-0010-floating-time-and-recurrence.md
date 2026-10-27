# ADR-0010: Floating local time, clamped monthly recurrence, skip exceptions
Status: **Accepted 2026-10-27 by the owner** (design sign-off)
Date: 2026-10-25

## Context
The calendar is for one household in one timezone. The prototype stored dates with no timezone,
skipped monthly events on days that do not exist, and had no way to cancel one occurrence.

## Decision
- Events are stored as local wall-clock dates and times in the calendar's timezone, with no
  conversion to UTC. The timezone decides only "today" and "now".
- Monthly events on the 29th to 31st **clamp** to the last day of shorter months; yearly 29
  February clamps to 28 February. (Prototype: skip.)
- Recurrence supports every N units and weekly weekday selection.
- A single occurrence can be skipped through `event_exceptions`; overriding one occurrence is
  reserved for later.

## Consequences
Good: no daylight-saving bugs for stored data; the rules are easy to explain and test; maps
cleanly to iCalendar for a later ICS import.
Bad: a later multi-timezone use case needs a per-event timezone; clamping differs from the
iCalendar default (skip), so an importer must translate such rules deliberately.

## Alternatives considered
UTC storage with per-event timezone (needless complexity for one household); skipping short
months (loses reminders); storing raw RRULE strings (powerful but opaque to the UI and harder to
validate).

## Addendum 2026-10-25
Monthly repeats may also be by weekday (first to fourth, or last, such weekday), selected under a
"More options" panel in the form. A multi-day event may carry a start time on its first day and an end
time on its last day (owner decided 2026-10-27).
