# 02 - Architecture

Status: **proposed**. The prototype implements most of this, but the design is meant to be
reviewed and changed before production code is written.

## 1. System context

```mermaid
flowchart LR
  user([You, in a browser])
  subgraph home["Home network"]
    app["Calendar app<br/>web UI + API + renderer + publisher"]
    db[("SQLite<br/>events, sections, settings")]
    terminus["Terminus server<br/>(BYOS)"]
    device["TRMNL device"]
  end
  user -->|"add events, choose layout"| app
  app --- db
  app -->|"1. login, create/update screen (image URL)"| terminus
  terminus -->|"2. GET /calendar.png"| app
  device -->|"3. check in on its refresh interval"| terminus
  terminus -->|"4. current screen image"| device
```

Key idea: the **app owns the data and the drawing**. Terminus is treated as a delivery
channel. This keeps layout logic out of Liquid templates and lets the app be developed and
tested without a device. (ADR-0002, ADR-0006)

## 2. Components inside the app

```mermaid
flowchart TB
  subgraph web["Web layer (Flask)"]
    ui["Editor UI<br/>static HTML + JS"]
    api_events["/api/events"]
    api_sections["/api/sections"]
    api_settings["/api/settings"]
    api_push["/api/push"]
    img["/calendar.png"]
  end
  subgraph domain["Domain"]
    validate["Validation"]
    recur["Recurrence expansion<br/>events_by_day()"]
  end
  subgraph render["Render engine"]
    canvas["Canvas<br/>fonts, 1-bit output"]
    layouts["Layout modules<br/>week | twoweek | month | agenda | split"]
    lanes["Section lane model<br/>weights, visibility, hide-empty"]
  end
  subgraph pub["Publisher"]
    worker["Background worker<br/>debounce + timer"]
    client["Terminus client<br/>login, POST/PATCH screen"]
  end
  store[("SQLite store")]

  ui --> api_events & api_sections & api_settings & api_push
  ui --> img
  api_events & api_sections & api_settings --> validate --> store
  api_events & api_sections & api_settings -. "request_push()" .-> worker
  api_push --> client
  worker --> client
  img --> layouts
  layouts --> lanes
  layouts --> canvas
  layouts --> recur --> store
  lanes --> store
```

Prototype note: in the prototype, `store.py` holds both storage and domain logic, and
`render.py` holds canvas, lanes and every layout. The split above is a **proposed**
separation; see `app/README.md` for the target folder structure.

## 3. Sequence: adding an event through to the device

```mermaid
sequenceDiagram
  actor U as User
  participant UI as Browser UI
  participant API as App API
  participant DB as SQLite
  participant W as Publisher worker
  participant T as Terminus
  participant D as TRMNL device

  U->>UI: Save event
  UI->>API: POST /api/events
  API->>DB: validate, then INSERT
  API-->>UI: 201 Created
  UI->>API: GET /calendar.png (refresh preview)
  API-->>UI: PNG
  API->>W: request_push("change")
  Note over W: wait 3 s to merge a burst of edits
  W->>T: POST /login
  T-->>W: access_token
  W->>T: PATCH /api/screens/{id} with the image URL
  T->>API: GET /calendar.png
  API->>DB: read events, sections, settings
  API-->>T: 1-bit PNG
  T-->>W: 200 OK
  Note over T,D: later, on the device's refresh interval
  D->>T: ask for current screen
  T-->>D: cached image
  D->>D: redraw e-paper
```

Consequence: a new event reaches the device only on its **next check-in**. The delay is
bounded by the device refresh interval (Terminus's example shows 900 s), not by this app.

## 4. Publisher worker loop

```mermaid
flowchart TD
  s([start]) --> d5["sleep 5 s"] --> p0["push('startup')"] --> w
  w{"wait for wake event<br/>timeout = push interval"}
  w -- "woken by a change" --> deb["sleep 3 s (debounce)"] --> clr["clear event"] --> pr["push(reason)"] --> w
  w -- "timeout" --> ps["push('schedule')"] --> w
```

`push()` never raises; it records `{state, message, at}` for the UI. First push creates the
screen and stores its id in the database; later pushes PATCH that id; a 404 recreates it.

## 5. Process and threading model

- The publisher is an in-process daemon thread with **in-memory status**. That forces **one
  application process** (for example `gunicorn --workers 1 --threads 4`). Two workers would
  push twice and show different statuses.
- Terminus calls back into the app while the push request is still open (it fetches
  `/calendar.png` during the PATCH), so the app needs **at least two request threads**.
- Alternative to evaluate: a separate publisher process or a cron-style job that calls
  `/api/push`, which removes the single-process constraint. (Open question OQ-7.)

## 6. Deployment (home network)

```mermaid
flowchart LR
  subgraph lan["LAN 192.168.x.x"]
    subgraph host1["Always-on machine A"]
      calendar["calendar app :8080"]
      vol[("volume: events.db")]
    end
    subgraph host2["Machine B (or same host)"]
      tm["Terminus + Postgres + Valkey"]
    end
    dev["TRMNL device (Wi-Fi)"]
  end
  calendar --- vol
  calendar <-->|HTTP| tm
  dev -->|HTTP| tm
```

- The image address Terminus uses (`TERMINUS_IMAGE_URL`) must be reachable **from Terminus**,
  which may differ from the address your browser uses. If Terminus runs in Docker, `localhost`
  is wrong.
- Give both machines DHCP reservations so addresses do not change.
- Do not expose the app to the internet: it has no authentication.

## 7. Extensibility points (the "modular" requirement)

| Extension point | Today | Proposed |
|-----------------|-------|----------|
| New layout | Add a branch in `render_png()` | Layout modules registered by name, each implementing `render(ctx) -> Canvas`. See `04-rendering-design.md`. |
| New section behaviour | Fields on `sections` table | Same table; add columns via migrations. |
| New setting | Add key to `DEFAULT_SETTINGS` + choice list | A declarative settings schema (key, type, default, choices, description) that drives validation, the API and the UI form. |
| New delivery target | Not possible | A `Publisher` interface; Terminus is one implementation (could add "write PNG to disk", "webhook"). |
| New event source | Manual entry only | An importer interface (ICS first). |

## 8. Alternative architecture to compare (not built)

**Terminus-native extension** (see `05-terminus-integration.md`, option B): the app only stores
and serves JSON (`/api/upcoming`); Terminus polls it on a schedule and renders a Liquid
template. Removes the push step and the Pillow renderer, but moves all layout work into
Liquid/HTML and couples the project to Terminus's extension format.
