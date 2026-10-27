# ADR-0007: SQLite, one application process, publisher state in the database
Status: **Accepted 2026-10-27 by the owner** (design sign-off)
Date: 2026-10-20

## Context
Home network, one user, small data. The first draft kept publisher status in memory, which
tied the app to a single process. The spike showed Terminus calls back into the app during a
push, so the app must stay responsive while a push is in flight.

## Decision
Use SQLite (WAL mode). Run one application process. Keep push state (last hash, time, status,
error, Terminus screen id) in a table. Startup asserts a single worker. Push runs off the
request loop; the image route serves from a render cache. Closes OQ-7. (The lease proposed in
the first revision was dropped as unnecessary, owner agreed 2026-10-25.)

## Consequences
Good: zero-admin storage; one file to back up; status survives restarts.
Bad: nothing stops a second instance except the startup check; SQLite writers are serial
(fine at this scale).

## Alternatives considered
PostgreSQL (needless here); a separate publisher process or cron job calling `/api/push`
(more parts to deploy and monitor).
