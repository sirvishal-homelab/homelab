# ADR-0013: A Screen page of its own

- Status: Accepted
- Date: 2026-11-01

## Context

Screen status, battery, Push now, the pushing switch and the connection check lived at the bottom of
Settings, below the calendar name and timezone. They are what the owner looks at most after the calendar
itself, and the design mockups gave them a tab.

## Decision

- `GET /screen` shows the picture on the screen, the last push, the battery, the device check-in, and the
  connection check. The forms post to `/screen/check`, `/screen/push` and `/screen/enabled` (they were
  `/settings/screen/...`).
- Settings keeps the household settings and the Look control, with a link to the Screen page.
- Phone tabs become Calendar, Month, **Screen**, More. Views moves under More. The wide top bar lists
  every page.
- The JSON API (`/api/v1/display/...`) is unchanged.

## Consequences

- Old `/settings/screen/...` form posts and the `#screen` anchor stop working. They were only used by
  the app's own pages.
- `NO_PUSH_PATHS` lists the new paths.

## Alternatives considered

- **Keep it in Settings and add a link:** least change, but the page stays long and the most used part
  stays at the bottom.
