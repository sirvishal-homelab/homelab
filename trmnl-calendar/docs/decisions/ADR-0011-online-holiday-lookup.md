# ADR-0011: Optional online holiday lookup
Status: **Proposed 2026-10-29** (built on the owner's request; the live service is not yet verified)
Date: 2026-10-29

## Context
Holiday regions are rule sets in code (`domain/holidays.py`): offline, tested against an official
chart, but each region is a code change. The owner does not need more regions. He wants the project to
be flexible for other people who reuse it.

## Decision
- Keep the built-in rules as the default. Add an **optional online lookup** page (`/holidays/online`,
  `GET/POST /api/v1/holidays/online...`) that fetches a country's public holidays, with its provinces or
  states, from Nager.Date (MIT licence, free, no key). Results go through the same picker and import as
  the built-in regions, so they become ordinary events with duplicate skipping.
- Only `holiday_source.py` talks to the service. It uses the standard library, a short timeout, a
  1 MB cap and an in-memory cache (1 hour for lists, 1 day for countries).
- The lookup runs **only when someone presses Look up** (or imports). It never runs on a schedule or on
  page load, and sends only a country code and a year. `HOLIDAY_LOOKUP=off` hides it;
  `HOLIDAY_API_URL` changes the address (infrastructure, so an environment variable, ADR-0005).
- The parser accepts the service's documented v3 fields and the announced rename `counties` to
  `subdivisionCodes`, and skips unusable entries.

## Consequences
Good: any country works without a code change; one small module; nothing changes for people who do not
use it; once imported, nothing depends on the service.
Bad: needs internet at import time; dates are only as good as the service (the built-in Nova Scotia
list stays the checked one); the service says v3 is supported until **2027-01-31**, so the module needs
a move to v4 once its reply is confirmed; this was written without being able to call the live service,
so the first real lookup is the verification (docs/08).

## Alternatives considered
More regions in code (best accuracy, one PR each; still the way to add a checked region); a UI to
define holiday rules (no outside dependency, but a rule editor to build and test); importing an .ics
file (reuses a future ICS importer, depends on the file's quality).
