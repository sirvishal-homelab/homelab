# ADR-0003: Push the image to Terminus through its Screens API
Status: Proposed (unverified; depends on spike T-1 to T-6)
Date: 2026-10-20

## Context
From the Terminus docs reviewed, a screen's image URI is fetched when the screen is created or
updated; no timed re-fetch was found. Extensions poll data, not images.

## Decision
The app logs in, creates one screen, remembers its id, and PATCHes it when data or layout
changes and on a timer (default 15 minutes).

## Consequences
Good: reuses the renderer; the app controls freshness; Terminus needs no custom code.
Bad: couples to Terminus's API and auth; needs a background worker (hence ADR-0007);
frequent logins; failure modes to design (see `05`).

## Alternatives considered
Option B (extension), option C (app speaks the device protocol directly). See `05`.
