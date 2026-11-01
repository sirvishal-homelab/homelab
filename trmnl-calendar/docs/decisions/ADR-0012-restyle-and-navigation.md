# ADR-0012: A modern look, a bottom bar on phones, and a bundled font

- Status: Accepted (tab list amended by ADR-0013)
- Date: 2026-11-01

## Context

The web pages were plain: a thin top bar, squared boxes, the system font. The owner reviewed design
mockups (cards, pill buttons, DM Sans, light and dark) and chose to restyle first and add new pages
afterwards. The app is self-hosted on a LAN and must keep working offline.

## Decision

- **Colours and shapes** are CSS variables in `app.css` (light, dark for Auto, dark for an explicit
  Dark). Cards have an 18 px radius and a soft shadow; buttons and chips are pills; targets stay at
  least 44 px.
- **Navigation.** Under 800 px wide the main links become a bar along the bottom (Calendar, Month,
  Views, More). Categories, People and Settings move under a new **More** page (`GET /more`), which also
  links to Who are you, Recently deleted and the data download. From 800 px up the top bar keeps every
  link. One `<nav>` serves both; links are marked `only-wide` or `only-narrow`.
- **Home** shows the screen picture beside "Coming up" from 1000 px up, and one column below that.
- **Font.** DM Sans (Regular, Medium, Bold; Latin subset, WOFF) is served from `/static/fonts` with its
  SIL Open Font License text. No request goes to a font service, so it works offline. The system font is
  the fallback.

## Consequences

- No new data or API. One new page. The e-paper picture is unchanged.
- About 85 KB of fonts are added to the image; they are cached by the browser.
- A **Shifts** page and a **Screen** page, as in the mockups, are a later slice; until then shift patterns
  stay in the event form and the screen status stays in Settings.
- Keep the two dark blocks in `app.css` in step.

## Alternatives considered

- **Load DM Sans from Google Fonts:** simpler, but it needs the internet and sends a request to a third
  party for every page view. Rejected for a LAN-only app.
- **A bottom bar on every screen size:** looks odd on a wide screen. Rejected; the top bar stays there.
- **Seven tabs on a phone:** too small to tap. Rejected for four tabs plus More.
