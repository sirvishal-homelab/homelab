# ADR-0003: Settings in SQLite, edited on a small web page
Status: Proposed
Date: 2026-11-14

## Context
The owner wants to choose which systems are shown without editing files. TRMNL Calendar
already keeps configuration in SQLite with a server-rendered page (its ADR-0005, ADR-0007,
ADR-0009).

## Decision
One process: FastAPI + Jinja pages, SQLite in a Docker volume, one background loop.
Secrets stay in the environment; everything else is in the database.

## Consequences (good and bad)
- Good: changes from a phone; survives rebuilds via the volume; familiar stack.
- Bad: a small web app to maintain; no login (LAN only, holds no secrets).

## Alternatives considered
- YAML config file: simpler, but every change needs SSH and a restart.
