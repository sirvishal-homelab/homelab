# ADR-0002: Separate calendar app that owns data and drawing; Terminus is delivery
Status: Proposed
Date: 2026-10-20

## Context
Terminus can render its own HTML/Liquid screens through extensions, and it can display images.
The user wants custom layouts (week, two-week, month, sections) and a place to enter events.

## Decision
Build a standalone app with its own database, editor and renderer. Terminus only receives the
finished image.

## Consequences
Good: layouts testable without a device; no dependency on Terminus internals beyond a small
API; app can serve other targets later.
Bad: an extra service to run; delivery depends on the Screens API (ADR-0003).

## Alternatives considered
Terminus extension that polls JSON and draws with Liquid (option B in `05`).
