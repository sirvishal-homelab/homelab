# 01 - Requirements

Status column refers to the prototype in `reference/prototype/`:
**Done** = exists and was exercised, **Partial** = exists but untested or not wired,
**None** = not started.

## Functional

| ID | Requirement | Prototype |
|----|-------------|-----------|
| FR-1 | Add, edit and delete events through a web UI. | Done (v1 UI) |
| FR-2 | Event fields: title, date, optional end date (multi-day), all-day or start/end time, location, repeat rule, section. | Done except section in UI |
| FR-3 | Repeat rules: none, daily, weekly, monthly, yearly, optional end date. | Done |
| FR-4 | Events are stored in a database and survive restarts and upgrades. | Done (SQLite) |
| FR-5 | Render the calendar as a black-and-white PNG sized for the device. | Done |
| FR-6 | Selectable layouts: week, two-week, month (and agenda variants). | Partial (renderer yes, no UI) |
| FR-7 | User-defined **sections** inside each day cell: create, rename, reorder, set relative height, hide, delete. | Partial (DB + renderer yes, no UI or API) |
| FR-8 | Each event belongs to one section; deleting a section moves its events to the first remaining section. | Partial (DB logic yes) |
| FR-9 | Layout options stored in the database: default view, week start, 12/24-hour, show weekends, show section labels, hide empty sections, show "updated" stamp. | Partial (DB yes, no UI or API) |
| FR-10 | Push the rendered image to Terminus when data or layout changes and on a timer. | Partial (code written, untested, not wired) |
| FR-11 | Show push status in the UI and allow "push now". | None |
| FR-12 | Preview exactly what the device will show, inside the editor. | Done (v1) |
| FR-13 | Machine-readable list of upcoming occurrences. | Done (`/api/upcoming`) |

## Non-functional

| ID | Requirement |
|----|-------------|
| NFR-1 | Runs on a small home server or Raspberry Pi. Prefer few moving parts. |
| NFR-2 | Configuration is **modular**: layouts, sections and settings are independent, data-driven pieces; adding a layout should not require touching unrelated code. |
| NFR-3 | Secrets (Terminus password) are not stored in the database. They come from the environment. |
| NFR-4 | No login in v1, so it must be safe to run on a trusted LAN only; document this. |
| NFR-5 | Images are legible on 1-bit e-paper: no thin white-on-black small text, no reliance on grey. |
| NFR-6 | Upgrades migrate the database in place without data loss. |
| NFR-7 | Rendering is deterministic for a given (data, settings, "now"), which makes it testable. |
| NFR-8 | The Terminus integration tolerates Terminus being down and recovers on the next push. |

## Explicitly undecided

See `08-open-questions-and-risks.md` (device model and bit depth, auth, ICS import,
per-occurrence exceptions to repeating events, how many sections, timezone handling).
