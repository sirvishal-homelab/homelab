# 06 - Application HTTP API

Status: **Proposed, revised 2026-10-27.** Replaces the prototype-era list. The API serves the browser UI
first; nothing outside the app depends on it (OQ-10: UI-only, but versioned so tests can rely on it). The
framework generates an OpenAPI description from the code; that description is the contract the API tests
check (`18`).

> **Implemented so far (slice 1):** events (create, read, list, patch with `If-Match`, delete, restore, skips),
> `GET /occurrences`, read-only `GET /categories`, `/healthz`, `/render/current.png`, `/render/{hash}.png`, and
> the single error shape. Handlers take plain JSON bodies, so the generated OpenAPI description does not yet
> describe request fields; typed models and the OpenAPI contract tests come with slice 2. `people` is always
> empty until slice 2. Not yet built: people, category editing, views, display and push endpoints.

## 1. Conventions

| Topic | Rule |
|-------|------|
| Base path | `/api/v1`. Pages and images sit outside it (section 4). |
| Format | JSON, UTF-8. Dates `YYYY-MM-DD`, times `HH:MM` 24-hour, timestamps UTC ISO-8601 with `Z` (`03`). |
| Updates | **PATCH with only the fields to change.** No full-replace `PUT` for events (closes the prototype's reset-on-omit trap). |
| Errors | One shape for every failure (section 5) with a human message the UI can show as is. |
| Auth | None in v1; LAN only. A "who am I" cookie selects a person for defaults and attribution; it is **not** authentication. A middleware hook is reserved for an optional shared password. |
| Push | Every call that changes events, categories, people, views, the calendar or the display asks the publisher for a debounced push. |
| Idempotent creates | `POST` creates accept an `Idempotency-Key` header. A repeat with the same key within 24 hours returns the original response and creates nothing (protects against double taps and retries after a dropped reply; `15`, M-6). |
| Edit conflicts | An event carries `rev`. `PATCH` and `DELETE` send `If-Match: <rev>`. A stale value returns **409** with the current event, who changed it and when. A missing header returns **428**. |
| Soft delete | `DELETE` marks an event deleted (30-day recovery); not a hard delete (`03`). |

## 2. Resources

### Events

| Method and path | Purpose |
|-----------------|---------|
| `GET /events?from=&to=` | Event series that overlap the window (not expanded). Optional `category`, `person`, `q`. |
| `GET /events?deleted=true` | Recently deleted events (last 30 days). |
| `GET /events/{id}` | One event, including `repeat_text`. |
| `POST /events` | Create. 201 and the event. Accepts `Idempotency-Key`. |
| `PATCH /events/{id}` | Partial update. Requires `If-Match`. Returns the event with the next `rev`. |
| `DELETE /events/{id}` | Soft delete. Requires `If-Match`. Returns 200 with the event so the UI can offer Undo. |
| `POST /events/{id}/restore` | Undo a delete within 30 days. |
| `POST /events/{id}/skips` | Body `{ "date": "2026-11-03" }`: skip one occurrence. |
| `DELETE /events/{id}/skips/{date}` | Bring a skipped occurrence back. |
| `GET /occurrences?from=&to=` | Expanded occurrences for the phone glance view (window up to 366 days). Optional `category`, `person`. |

Event body (update takes any subset):

```json
{
  "id": 41, "uid": "7c1e...", "rev": 3,
  "title": "Skating lesson",
  "category_id": 1,
  "start_date": "2026-11-07", "end_date": null,
  "start_time": "10:00", "end_time": "11:00",
  "location": "Arena",
  "repeat": { "freq": "weekly", "interval": 1, "weekdays": ["SA"],
              "month_mode": null, "month_nth": null, "until": "2027-03-27" },
  "repeat_text": "Repeats every week on Saturday, until 27 March 2027",
  "people": { "ids": [2], "owner_id": 2 },
  "updated_at": "2026-10-27T21:40:11Z", "updated_by": { "id": 1, "label": "A" },
  "deleted_at": null
}
```
`people.ids` empty means everyone. `repeat_text` is read-only, built by the server so the form can describe
the rule in words.

### Categories (sections), people, calendar

| Method and path | Purpose |
|-----------------|---------|
| `GET/POST /categories` | List in display order; create `{name, weight?, visible?, placement?}` (max 8). |
| `PATCH /categories/{id}` | Partial update. |
| `DELETE /categories/{id}?move_to=<id>` | Delete; events move to `move_to` (required if it has events). Refuses the last category. |
| `PUT /categories/order` | Body `{ "ids": [...] }`, a permutation of existing ids. |
| `GET/POST /people` | List; create `{name, label}` (max 10). |
| `PATCH /people/{id}` | Rename, relabel, `active: false` to retire. |
| `DELETE /people/{id}` | Only if no event uses the person (409 otherwise; retire instead). |
| `GET/PATCH /calendar` | Name, timezone, `week_start`, `time_format`. |
| `GET /me`, `PUT /me` | Read or set "who am I" (`{person_id}`); stored in a cookie. |

### Views, layouts, templates

| Method and path | Purpose |
|-----------------|---------|
| `GET /layouts` | Registered layouts with id, label and an options schema that drives the form (`04`). |
| `GET /templates` | Built-in templates (Week - lanes, Two weeks - list, Blank) as starting values. |
| `GET/POST /views` | List; create from a template: `{template, name, overrides?}`. |
| `GET/PATCH/DELETE /views/{id}` | Read, partial update, delete (refuses the active view). |
| `PUT /views/order` | Reorder. |
| `POST /views/{id}/activate` | Make it the display's active view and request a push. |
| `POST /views/preview` | Body `{ "view": {...unsaved settings...}, "as_of": "2026-10-25"? }`. Renders with the same code as the push; **stores nothing, pushes nothing**. Returns `{ "image": {"url": "/render/<hash>.png", "width": 800, "height": 480, "bit_depth": 2}, "fit": [ {code, message, where} ] }`. |

The preview image stays in the render cache for a few minutes and is fetched by its hash URL.

### Display and publishing

| Method and path | Purpose |
|-----------------|---------|
| `GET /display` | Profile (model name, size, depth), active view, enabled. |
| `PATCH /display` | Override depth or size; enable or disable pushing. |
| `GET /display/status` | `push_state` plus next check time, Terminus reachability, last error. No secrets. |
| `POST /display/push` | Push now, even if unchanged. **502** with details if Terminus fails. |

### Data and system

| Method and path | Purpose |
|-----------------|---------|
| `GET /export` | The whole configuration and all events as JSON (also a safety net). Import comes later. |
| `GET /healthz` | Cheap, unauthenticated: database reachable, scheduler alive. For Docker and Beszel. Outside `/api/v1`. |
| `GET /version` | App version and schema version. |

## 3. Key flows

**Add an event (phone).** `POST /events` with a fresh `Idempotency-Key`; the reply includes the event and
the time the screen should show it ("on the screen by about 3:15 PM", from `GET /display/status` and the
15-minute check-in). A failed or slow reply is retried with the same key: never a duplicate.

**Two people edit the same event.** B opens rev 3; A saves (rev 4). B's `PATCH` with `If-Match: 3` returns
409 with the current event and "changed by A at 3:02 PM"; the UI offers to review or keep B's edit (resent
against rev 4).

**Delete and undo.** `DELETE` returns the event; the UI shows Undo for about ten seconds and a "Recently
deleted" list for 30 days; Undo calls `POST /events/{id}/restore`.

**Preview while editing a view.** The UI debounces (about 300 ms), calls `POST /views/preview` with unsaved
settings and shows the image at 1:1 or fit. Nothing is saved until "Save view".

## 4. Pages and images outside `/api/v1`

| Path | Purpose |
|------|---------|
| `GET /` and friends | Server-rendered pages (`07`, `17`). |
| `GET /render/{hash}.png` | The image Terminus fetches during a push. Content-addressed, immutable, served from the render cache; never re-renders. |
| `GET /calendar.png` | Stable URL for the latest render of the active view. Used only by experiment E-1 (ADR-0003). |
| `GET /static/...` | Local CSS, script and vendored UI library (no CDN; works offline on the LAN). |

## 5. Errors

```json
{ "error": { "code": "validation_failed",
             "message": "Please check the highlighted fields.",
             "fields": { "title": "Enter a short description" } } }
```

| Status | Code | When |
|--------|------|------|
| 400 | `bad_request` | Malformed JSON, unknown query value |
| 404 | `not_found` | Unknown id |
| 409 | `edit_conflict` | Stale `If-Match`; body carries `current`, `changed_by`, `changed_at` |
| 409 | `in_use` | Deleting a person or category still referenced, without a target |
| 422 | `validation_failed` | Field errors, in plain language |
| 428 | `precondition_required` | `If-Match` missing on `PATCH` or `DELETE` |
| 502 | `terminus_error` | Terminus unreachable or rejected a push |

## 6. Not in v1

Authentication, import (export exists), a third-party API, named profiles for several devices, pagination
(windows bound every list), bulk edit.

## 7. Decisions this closes

OQ-10 (UI-only but versioned; PATCH not PUT) and prototype defects 3 and 5 in `08`.
