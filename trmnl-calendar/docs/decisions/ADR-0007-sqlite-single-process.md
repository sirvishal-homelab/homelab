# ADR-0007: SQLite and a single application process
Status: Proposed
Date: 2026-10-20

## Context
Home-network, single-user, small data. The publisher runs as a thread with in-memory status.

## Decision
Use SQLite. Run one application process with several threads. Revisit if the publisher moves
out of process.

## Consequences
Good: zero-admin storage; one file to back up.
Bad: the thread model ties the app to one process; concurrent writers are limited (fine here).

## Alternatives considered
PostgreSQL (needless for this scale); separate publisher process or scheduled job (removes the
single-process limit; open question OQ-7).
