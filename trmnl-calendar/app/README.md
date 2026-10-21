# app/ - the calendar application (FUTURE)

Empty on purpose. Production code starts here only after `docs/` is agreed
(see `docs/10-roadmap.md`). Everything below is a **proposal** to review, including the
language: Python is what the prototype used, not a decision (OQ-11).

## Proposed structure

```
app/
├── README.md
├── pyproject.toml                 dependencies, tooling (format, lint, test)
├── Dockerfile                     single process, bundled font, non-root user
├── docker-compose.yml             app + volume; optional Terminus example
├── .env.example                   every environment variable, documented
├── fonts/                         bundled font files and their licences
├── migrations/                    numbered, forward-only SQL (0001_init.sql, ...)
├── src/trmnl_calendar/
│   ├── config.py                  env parsing; secrets and infrastructure only
│   ├── domain/
│   │   ├── events.py              Event model, validation
│   │   ├── recurrence.py          expansion, exceptions (if approved)
│   │   ├── sections.py            Section model, lane rules
│   │   └── settings.py            declarative settings schema (key, type, default, choices)
│   ├── storage/
│   │   ├── db.py                  connection, migration runner
│   │   └── repositories.py        events, sections, settings access
│   ├── render/
│   │   ├── canvas.py              fonts, scale, 1-bit output
│   │   ├── lanes.py               weights, visibility, hide-empty
│   │   ├── registry.py            layout registration by name
│   │   └── layouts/               week.py, twoweek.py, month.py, agenda.py, split.py
│   ├── publish/
│   │   ├── base.py                Publisher interface
│   │   ├── terminus.py            Screens API client
│   │   └── worker.py              debounce + schedule, status
│   └── web/
│       ├── app.py                 app factory
│       ├── routes_events.py
│       ├── routes_sections.py
│       ├── routes_settings.py
│       ├── routes_push.py
│       ├── routes_image.py
│       ├── templates/
│       └── static/
└── tests/
    ├── fixtures/                  fixed events, sections, settings
    ├── snapshots/                 expected PNGs for a fixed "now"
    ├── test_recurrence.py
    ├── test_validation.py
    ├── test_render_snapshots.py
    ├── test_api.py
    └── test_terminus_client.py    against a fake Terminus
```

## Rules this structure is trying to enforce

1. `domain/` imports nothing from `web/`, `render/` or `publish/`.
2. `render/` takes plain data in and returns an image; it does not touch the database.
3. `publish/terminus.py` is the only file that knows Terminus's API.
4. Adding a layout means adding a file in `render/layouts/` and registering it.
5. Settings are declared once (`domain/settings.py`) and drive validation, API and UI.
6. Environment variables never hold user-editable layout choices; the database does.
