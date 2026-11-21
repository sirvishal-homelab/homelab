# ADR-0004: Pending-updates dot from Beszel's package-update counts
Status: **Accepted 2026-11-21 by the owner** (asked for it in chat and to implement it)
Date: 2026-11-21

## Context
The owner wants to see at a glance when a focus system has OS updates waiting. Beszel 0.21.0
added a pending-updates check to the agent: each system record carries `info.pu = [total,
security]`, refreshed hourly, and Beszel's own All Systems table shows it as a coloured dot.
The app already reads that record (ADR-0002), so no new data source is needed.

## Decision
Draw a small dot after a focus system's name: red if `security > 0`, amber if `total > 0`,
nothing otherwise. Same colours as Beszel's Updates column, without the green "up to date".
Only on rows that are up (a down or paused system's counts are stale). Never an alert card:
updates are routine, not an outage. One global switch on the settings page, on by default
(`show_updates` in `settings`, no migration).

`info.pu` is optional: missing on agents running in Docker or inside containers. Missing or
malformed values mean "not reported" (no dot), not "Beszel data changed" (F-16), because a
cosmetic field must not be able to take over the whole screen.

## Consequences (good and bad)
- Good: no new credentials, endpoints or agents; matches what the owner sees in Beszel.
- Good: hourly refresh is fine for a glance; the frame only changes when the dot does.
- Bad: systems whose agent runs in Docker show no dot at all (indistinguishable from "up to date"
  on the screen). Installing the agent as a system service fixes that if wanted.
- Bad: a renamed or reshaped `pu` field after a Beszel update hides the dots silently. The
  contract tests against the recorded fixture catch it when the fixtures are refreshed.

## Alternatives considered
- Proxmox API (`/nodes/{node}/apt/update`): hosts only, needs a token; breaks ADR-0002.
- Per-guest script or SSH from the app: covers containers too, but setup on every system.
- Per-system switch: more settings for little gain; the global switch is enough for now.
