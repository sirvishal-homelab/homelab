# ADR-0006: Render server-side with Pillow to a 1-bit PNG
Status: Proposed
Date: 2026-10-20

## Context
The display is 1-bit. Layouts must be crisp and deterministic and run on a small machine.

## Decision
Draw on a grey canvas with Pillow, then threshold to pure black and white and save a 1-bit PNG.
Layouts are modules registered by name (proposed interface in `04`).

## Consequences
Good: light, fast, no browser; deterministic and snapshot-testable.
Bad: manual layout math and text wrapping; a bundled font is required; no CSS.

## Alternatives considered
Headless browser screenshot of HTML (richer layout, heavy and slower); Terminus-side Liquid
rendering (ADR-0002 alternative).
