# 18 - Test strategy

Status: **Proposed, 2026-10-27.** Written before any code. Test rules where they live, keep most tests fast
and pure, and be honest about what only a person can judge.

## 1. Layers

| Layer | What it proves | Tooling | Runs |
|-------|----------------|---------|------|
| Unit: domain | Validation; **recurrence** (daily, weekly with interval and weekdays, monthly by date and by weekday, yearly, clamping, leap years, month ends, `until`, skips, multi-day spans, window edges); filters and ordering | pytest, data builders | every commit |
| Unit: render rules | Scene snapshots (JSON); no light-on-dark text; no text under 10 px; text inside cells; each fit finding triggered; eight categories; 5- and 7-day; today hidden; trips across rows; person filter | pytest | every commit |
| Format | PNG header colour type 0 with the profile's depth and 800 x 480; only allowed grey levels; same inputs give the same hash; font checksum | pytest | every commit |
| Image goldens | PNG per template and fixture is pixel-exact | pytest inside the container image | every pull request |
| Storage | Constraints, soft delete and purge, idempotency, backups, restore; **migrations from every earlier schema** using fixture database files | pytest | every pull request |
| API | Every endpoint against the generated OpenAPI contract; error shape; `If-Match` 409 and 428; idempotent replays; `PATCH` leaves omitted fields alone | pytest, httpx; optionally Schemathesis | every pull request |
| Terminus contract | The client against a **fake Terminus** (section 3) | pytest | every pull request |
| Scheduler | Debounce, 15-minute check, push-only-on-change, retry and backoff, restart recovery | pytest, fake clock | every commit |
| End to end | Phone (390 x 844) and desktop: add an event in a few taps, undo a delete, conflict warning, draft survives a failed save, view editor with live preview, accessibility checks | Playwright | CI and before release |
| Manual | The physical panel; real Terminus; the wife's trial | checklists, `terminus-spike.py` | before a release or a Terminus upgrade |

## 2. Fixtures and determinism

- **Fixed now:** 2026-10-25 09:00 America/Halifax (a Sunday, the day the options were designed).
- **Sample dataset:** the events behind the layout options in `16` and `17` (trips, holidays, two
  placeholder people). Goldens for the two first templates should match the approved mock-ups in
  `assets/options/`; differences are reviewed, never silently accepted.
- **Builders**, not hand-written rows. **Per-version database files** back the migration tests.
- **Stable rendering:** the image pins Python, Pillow and FreeType; fonts are bundled and hash-checked.
  Goldens run **inside that image in CI** (a different FreeType changes glyph shapes) and are regenerated
  only by an explicit, reviewed command.

## 3. The fake Terminus

A small HTTP server in `tests/contract/fake_terminus/` implementing only what the spike verified (`14`):

| Behaviour | Spike item |
|-----------|-----------|
| `POST /login` returns access and refresh tokens; `POST /api/jwt` refreshes | T-3, T-9 |
| Header is the bare token | T-3 |
| `GET /api/models` and `/api/devices` with the observed fields | T-4 |
| `POST` and `PATCH /api/screens` return **200**; the server **fetches the URI during the call** and stores a `preprocessed` PNG unchanged | T-1, T-5, T-6, T-8 |
| A PATCH keeps the screen id and changes its stored `uri` | T-8 |

Fault injection: login failure, expired token, 401 mid-session, unknown screen (404), slow response, 5xx, and
a callback that cannot reach the app. Before any Terminus upgrade, `terminus-spike.py` runs against the real
instance; if output differs from `docs/spikes/`, the contract and the fake change together.

## 4. Continuous integration (stage 1 of ADR-0008)

Per pull request: lint and format check, type check, unit, format, storage, API, contract and scheduler
tests, image goldens in the container, Docker build, start-up smoke test (`/healthz`), Playwright. Merge to
`main` is blocked unless green. Playwright may be split into a quick smoke set per PR and a full set before
a release.

## 5. What automation cannot judge

| Item | How it is checked |
|------|-------------------|
| Readability and grey contrast on the panel | Push goldens with `terminus-spike.py --push-file` and look; checklist per template |
| Real Terminus after an upgrade | Run the spike script and compare with the log |
| "Easy for her" | Scripted trial with the wife on her phone before each slice is done: add an event, add a trip, make a repeating event, delete and undo; note every hesitation |
| Anti-aliasing of larger headings (RD-3) | One on-device test in the first rendering slice |

## 5a. Status after slice 1

Built: unit tests for recurrence (leap years, month ends, nth weekday, interval, `until`, skips, trips),
validation, repeat text, forms, the PNG encoder (round trip through Pillow at every bit depth) and the
Week - lanes rules (hidden days, today highlight, overflow, lane size, banners, trips, determinism, text rules);
integration tests for storage, migrations, the repositories, the render service, and the API and pages through
FastAPI's test client. **Not yet built:** pixel-exact image goldens (they must be generated inside the project's
container image, section 2), contract tests against the fake Terminus (slice 4), end-to-end browser tests.

## 6. Rules of thumb

1. A bug found by hand gets a failing test first.
2. A new layout needs a scene snapshot, a golden, rule tests and a fit-finding case.
3. A schema change needs a migration, a fixture database at the new version, and a migration test from every
   earlier version.
4. Tests never use the real clock, the network or the real Terminus.
