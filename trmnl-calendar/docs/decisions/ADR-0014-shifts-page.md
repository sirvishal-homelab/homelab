# ADR-0014: Shifts pages for rotating schedules

- Status: **Accepted** (owner, 2026-11-01)
- Date: 2026-11-01

## Context

Rotating-week schedules, with per-day In / Out and days off, were added on 2026-10-29. They are managed
from the event form and from buttons under each scheduled day in the home list. There is no single place
to see the pattern for the coming weeks, and no view of the work schedule alone. The owner's design
mockups show a Shifts page, and the owner asked for it to be its own tab, with a dedicated month view of
the work schedule that a partner can read at a glance and change one date at a time.

## Decision

Add a **Shifts** tab (always shown) with two pages: **Weeks** (`GET /shifts`, a card per schedule with a
strip as long as the rotation) and **Month** (`GET /shifts/month`, one schedule per month grid, with the
day's In / Out / Day off / Restore buttons below). Both are built from existing data: **no migration and
no new API**. State is shown in words and shapes. Pattern editing stays in the event form. A free-text
note per date is **not** part of this slice (owner's choice); it can be added later as one small table.

## Consequences

- Small, low-risk slice: view builders, two pages, tests.
- Several places can change a day (home list, month page, Shifts) and all post to the same routes, so
  they cannot disagree.
- The phone bar now has five tabs: Calendar, Month, Shifts, Screen, More.
- Editing the pattern still takes one extra tap through the event form.

## Alternatives considered

- **Pattern editor on the page:** closer to the mockup, but it needs a new save route and conflict
  handling, and it duplicates the event form. Deferred.
- **A Shifts tab only when a schedule exists:** keeps the bar short for other households. Rejected by
  the owner, who wants it always visible.
- **A free-text note on each date:** useful ("training 8 to 4") but needs a table, export and restore
  changes. Deferred by the owner.
