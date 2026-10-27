# ADR-0006: Render server-side with Pillow to a grayscale PNG in two phases
Status: **Accepted 2026-10-27 by the owner** (design sign-off)
Date: 2026-10-20

## Context
The display is 800x480 and 2-bit on the owner's panel. The spike showed Terminus stores a
preprocessed PNG unchanged and the panel shows four greys. Layouts must be crisp, deterministic,
snapshot-testable, and previewable at native size.

## Decision
Layouts **plan** a scene of primitives (pure, no drawing) and a generic rasteriser **draws** it with
Pillow and a bundled font, quantises to the profile's depth, and encodes a grayscale PNG (colour
type 0, 1 or 2 bit) with a small own encoder. The same function serves the preview and the push.
Details and the module interface are in `04`.

## Consequences
Good: light, fast, no browser; deterministic; a fit check and JSON snapshot tests fall out of the
plan phase; one code path for preview and push.
Bad: manual layout math and text wrapping; a bundled font and pinned Pillow and FreeType are
required for stable snapshots; no CSS.

## Alternatives considered
Headless browser screenshot of HTML (richer layout, heavy and slower, and uncertain output); Terminus
side Liquid rendering (ADR-0002 alternative).
