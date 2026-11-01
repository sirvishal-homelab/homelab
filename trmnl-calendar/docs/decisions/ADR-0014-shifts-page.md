# ADR-0014: A Shifts page for rotating schedules

- Status: **Proposed** (owner to review `docs/20-shifts-page.md`)
- Date: 2026-11-01

## Context

Rotating-week schedules, with per-day In / Out and days off, were added on 2026-10-29. They are managed
from the event form and from buttons under each scheduled day in the home list. There is no single place
to see everyone's pattern for the coming weeks. The owner's design mockups show a Shifts page.

## Decision

Add `GET /shifts`, a read-and-quick-edit overview built from the existing data: one card per rotating
schedule, a strip as long as the rotation, state shown in words, and the existing In / Out / Day off /
Restore actions. No migration and no new API. Pattern editing stays in the event form for now. Shifts is
a phone tab only when a schedule exists.

## Consequences

- Small, low-risk slice: one pure domain function, one view builder, one page, tests.
- Two ways to change a day (home list and Shifts page) call the same routes, so they cannot disagree.
- Editing the pattern still takes one extra tap through the event form.

## Alternatives considered

- **Pattern editor on the page:** closer to the mockup, but it needs a new save route and conflict
  handling, and it duplicates the event form. Deferred.
- **Always show a Shifts tab:** a dead tab for households without shifts. Rejected.
