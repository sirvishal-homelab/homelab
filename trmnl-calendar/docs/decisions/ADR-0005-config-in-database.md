# ADR-0005: Configuration lives in the database; environment holds secrets and seeds defaults
Status: Proposed
Date: 2026-10-20

## Context
The user wants configuration to be modular and not lost between changes or restarts.

## Decision
Layout options and sections are stored in SQLite. Environment variables may seed defaults
(`WEEK_START`, `TIME_FORMAT`) but once a value is saved, the database wins. Secrets and
infrastructure addresses (`TERMINUS_*`, timezone, ports) stay in the environment.

## Consequences
Good: settings survive rebuilds; the UI can edit them; no secrets in the data file.
Bad: two places to look; the precedence rule must be documented; backup must include the DB.

## Alternatives considered
All settings in environment (not UI-editable); a config file (not transactional with data).
