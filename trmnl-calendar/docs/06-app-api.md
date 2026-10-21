# 06 - Application HTTP API

Conventions (prototype): JSON in and out. Validation failures return **400** with
`{"errors": {"<field>": "<message>"}}`. Other failures return `{"error": "<message>"}`.
Success on delete returns **204**. No authentication (LAN only).

## Existing in the prototype's `app.py` (v1, working)

| Method and path | Purpose |
|-----------------|---------|
| `GET /` | Editor page. |
| `GET /calendar.png` | The image. Query: `view`, `w` (200-2000, default 800), `h` (200-2000, default 480), `date=YYYY-MM-DD` (pretend "today", for previews and tests). Sent with `Cache-Control: no-store`. |
| `GET /api/events` | All events. |
| `POST /api/events` | Create. Returns 201 and the event. |
| `PUT /api/events/<id>` | Replace an event. Note: **full replace**, so omitted fields reset to defaults. |
| `DELETE /api/events/<id>` | Delete (removes every repeat). |
| `GET /api/upcoming?days=N` | Expanded occurrences with repeats applied (1-366 days, default 30). |
| `GET /healthz` | Liveness. |

Event JSON: `id, title, date, end_date, start_time, end_time, location, repeat, repeat_until, section_id`
(input also accepts `all_day: true`, which clears the times).

## Planned (store functions exist; routes and UI do not)

| Method and path | Purpose |
|-----------------|---------|
| `GET /api/settings` | Current layout settings (see `03-data-model.md`). |
| `PUT /api/settings` | Partial update; only known keys; validated; triggers a push. |
| `GET /api/sections` | Ordered sections. |
| `POST /api/sections` | Create `{name, weight?, visible?}` (max 8). |
| `PUT /api/sections/<id>` | Partial update of `name`, `weight`, `visible`. |
| `POST /api/sections/order` | Body `{"ids": [...]}`, must be a permutation of existing ids. |
| `DELETE /api/sections/<id>` | Re-homes events to the first remaining section; refuses if last. |
| `GET /api/push` | Publisher status: `{configured, missing[], state, message, at, screen_id, interval_seconds, ...}`. No secrets. |
| `POST /api/push` | Push now and return the resulting status. |

Every mutating call on events, sections or settings should call `publisher.request_push()`.

## Design questions for the API

- Should `PUT /api/events/<id>` become `PATCH` (partial) to avoid the reset-on-omit trap? (OQ-10)
- Should the image URL accept a **named profile** (for example `/calendar.png?profile=kitchen`)
  to serve several layouts to several devices? Not requested; would need a `profiles` table.
- Versioning (`/api/v1/...`) once anything outside the UI depends on it.
- Pagination for `/api/events` if the table grows.
