# app/ - the calendar application (structure settled 2026-10-27)

Design signed off 2026-10-27. **Slice 0 (foundations) is built**: `/healthz`, migrations, backups, font check,
Dockerfile, compose, CI. Everything else below is the agreed structure for later slices.
Stack: ADR-0009 (Python 3.12, FastAPI, SQLite, Pillow, HTMX). Paths are relative to `app/`.

```
app/
├── README.md
├── pyproject.toml                  dependencies, ruff and pytest settings
├── uv.lock                         committed lock file
├── Dockerfile                      multi-stage, non-root, bundled fonts, HEALTHCHECK
├── .dockerignore  .env.example     every environment variable, documented
├── deploy/
│   ├── compose.yml                 the calendar's compose project for the server (volume /data)
│   └── README.md                   pull, update, roll back, restore (manual; stage 1 of ADR-0008)
├── src/trmnl_calendar/
│   ├── main.py                     app factory, startup checks (fonts, single worker, migrations)
│   ├── config.py                   environment: infrastructure and secrets only
│   ├── domain/                     pure; imports nothing from the other packages
│   │   ├── models.py               Event, Category, Person, View, Occurrence (typed)
│   │   ├── validation.py           the rules in docs/03
│   │   ├── recurrence.py           occurrences(events, exceptions, window)
│   │   ├── filtering.py            weekday, category and person filters, ordering
│   │   └── settings.py             declarative option schemas (drive validation and forms)
│   ├── storage/
│   │   ├── db.py                   connection, pragmas (WAL, foreign keys, busy timeout)
│   │   ├── migrate.py              runner and pre-migration backup
│   │   ├── migrations/             0001_baseline.sql, 0002_...  (forward-only)
│   │   ├── repositories/           events categories people views push_state idempotency
│   │   └── backup.py               nightly copy, retention, restore helper
│   ├── render/                     pure; no database, no clock
│   │   ├── fonts/                  bundled DejaVu Sans Condensed + LICENSE + checksums
│   │   ├── profile.py  scene.py  text.py  raster.py  quantise.py  png.py  fit.py
│   │   ├── layouts/                registry.py week_lanes.py twoweek_list.py (more later)
│   │   └── templates.py            Week - lanes, Two weeks - list, Blank
│   ├── publish/
│   │   ├── base.py                 Publisher interface
│   │   ├── terminus.py             the only file that knows Terminus's API
│   │   ├── push.py                 render, hash, send, record
│   │   ├── scheduler.py            one loop: debounce, 15-minute check, startup
│   │   └── cache.py                render cache: hash to PNG, small LRU
│   └── web/
│       ├── api/                    routers: events categories people calendar views display system
│       ├── pages.py  errors.py  deps.py   (who-am-I cookie, idempotency, If-Match)
│       ├── templates/              Jinja pages (mobile first)
│       └── static/                 css, small script (drafts, preview), vendored htmx
└── tests/                          see docs/18-test-strategy.md
    ├── unit/  snapshots/  integration/  contract/fake_terminus/  e2e/  fixtures/
```

## Rules this structure enforces

1. `domain/` and `render/` import nothing from `web/`, `storage/` or `publish/`. A test checks it.
2. `render/` takes plain data and a fixed `now`; it never reads the clock or the database.
3. `publish/terminus.py` is the only code that knows Terminus's API.
4. A new layout is one file in `render/layouts/` plus registration; a new template is one entry in
   `templates.py`. Neither needs a migration, an API change or a UI change.
5. Options are declared once (`domain/settings.py`, each layout's schema) and drive validation, API, forms.
6. Environment variables hold infrastructure and secrets only; everything the user edits is in the database.

## Outside this folder

- `homelab/infra/terminus/`: the Terminus compose file, a secrets-free `.env.example` and a README
  (`docs/13-infrastructure.md`, P-9).
- `homelab/.github/workflows/`: CI, filtered to `trmnl-calendar/**` (`docs/12-delivery-pipeline.md`).

## Develop

```bash
cd trmnl-calendar/app
uv sync
uv run ruff check . && uv run ruff format --check .
uv run pytest
DATA_DIR=./data uv run uvicorn --factory trmnl_calendar.main:app_factory --port 8080
```
