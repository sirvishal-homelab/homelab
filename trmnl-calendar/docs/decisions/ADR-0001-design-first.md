# ADR-0001: Design first; the prototype is reference only
Status: **Accepted 2026-10-27 by the owner** (design sign-off)
Date: 2026-10-20

## Context
A prototype was built quickly in one session. Parts are tested, parts are not, and the app,
UI and container files lag behind the data layer and renderer. The user is moving the work to a
longer-lived project and wants architecture and design settled before more code.

## Decision
Treat `docs/` as the source of truth. Keep the prototype under `reference/prototype/`,
labelled with `STATUS.md`. New production code goes in `app/` and `plugin/` only after the
design is accepted.

## Consequences
Good: avoids building on unproven code; keeps lessons without inheriting structure.
Bad: some work (UI, API wiring) will be redone.

## Alternatives considered
Continue the prototype in place (faster, but locks in module-level globals, ad hoc migrations
and the Docker problems).
