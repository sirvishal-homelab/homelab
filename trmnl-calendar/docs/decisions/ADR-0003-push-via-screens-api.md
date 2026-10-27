# ADR-0003: Push the image to Terminus through its Screens API
Status: **Accepted 2026-10-25 by the owner, with a review gate** (spike 2026-10-24, see `docs/14`)
Date: 2026-10-20, revised 2026-10-25

## Context
The spike against the owner's Terminus (OG, 2-bit model `og_plus`) showed:
- `POST` and `PATCH /api/screens` accept an image URI on the LAN and fetch it synchronously.
- With `preprocessed: true` the PNG is stored unchanged, including 4-level grey; unprocessed
  mode hard-thresholds to 1-bit.
- Bare-token auth works; tokens refresh through `POST /api/jwt`.
- Embedding an image in HTML `content` produced a blank screen (not viable).

## Decision
The app renders a PNG matching the model's bit depth and pushes it with `preprocessed: true`.
It resolves the model id from the model name at run time, creates one screen once, remembers
its id, and PATCHes it when data or layout changes and on a timer (default 15 minutes). A 200
response means the image was fetched and stored. Option B (extension) and `plugin/` are not
built.

## Consequences
Good: reuses the renderer; the app controls thresholding and freshness; Terminus needs no
custom code; errors surface immediately.
Bad: couples to Terminus's API (beta; pin the version); needs a background worker (ADR-0007);
every PATCH writes a new upload file (Terminus removes the replaced one, S-3); the calendar should be the only item in its
playlist or the device will alternate screens (S-2, observed). The physical display was checked: a 2-bit PNG
renders correctly (S-1).

## Alternatives considered
Poll extension with an image source: Terminus's upstream tracker shows poll extensions accept
image sources, so Terminus could pull our PNG on its own schedule with no credentials and no
publisher. **Not tested** (T-7; `extensions.adoc` not read in full). Concern: it re-renders on
every poll, which likely gives the device a new image each cycle, and it cannot push
"only when changed". Upstream issue #384 (open, 2026-10) adds a race: Terminus deletes the
previous image at once on re-render, so a device that fetched the old URL just before gets a 404.
Push-on-change keeps that window rare. This risk applies to our PATCH too.

Option B (extension), option C (app speaks the device protocol), HTML content with an embedded
image (tested and rejected). See `05`.

## Review gate: experiment E-1 (pull-mode trial)
The owner accepts push for the MVP but wants to test the alternative before it sticks. After
slice 4 (push) the app also offers a stable `GET /calendar.png` (latest render), and a Terminus
poll extension with that image URL as its source is trialled against push. The app's
`Publisher` interface keeps both swappable.

Measure, for about a week each (push, then pull), against the same panel:

| Measure | Question |
|---------|----------|
| M1 fidelity | Is the image Terminus stores bit-identical to ours (reuse the spike's comparison)? |
| M2 redraw | With an unchanged calendar, does Terminus store a new file and does the panel redraw each cycle? |
| M3 latency | Time from saving an event to the panel changing (5 samples each) |
| M4 failures | Broken-image or 404 incidents (upstream issue #384), errors, retries |
| M5 operations | Credentials needed? Setup steps? Recovery after restoring Terminus from backup |

Decision rule: pull replaces push only if M1 is exact, M2 shows no churn, M3 is no worse than
push plus 5 minutes, and M4 shows no incidents. Otherwise push stays. The owner decides; the
outcome is recorded here as an addendum, or in a superseding ADR if pull wins.
