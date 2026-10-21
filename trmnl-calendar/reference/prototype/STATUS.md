# Prototype status (read before reusing anything)

This code was written in one working session. It is **reference material**, not a foundation
the design must follow.

| File | State | Notes |
|------|-------|-------|
| `store.py` | New, partly exercised | SQLite storage for events, sections, settings; validation; recurrence expansion; inline migration. The v1 subset (events, recurrence, validation) was tested through the API. The sections and settings functions were only exercised indirectly by the renderer and `seed_demo.py`. |
| `render.py` | New, visually checked | Pillow renderer: `week`, `twoweek`, `month` (with sections), plus `agenda` and `split`. Checked by eye at 800x480 for several settings. No automated tests. |
| `publisher.py` | New, **untested** | Terminus push client (login, create screen once, PATCH afterwards) with a debounced background worker. Never run, not even against a mock server. |
| `seed_demo.py` | New, works | Sample events across sections. |
| `app.py` | **Stale (v1)** | Does not use sections, settings or the publisher. No `/api/settings`, `/api/sections` or `/api/push` routes. |
| `templates/`, `static/` | **Stale (v1)** | Event form and list only. No section manager, layout settings, section field on the event form, or Terminus status panel. Tested in a headless browser (v1 only). |
| `Dockerfile`, `docker-compose.yml` | **Stale and partly wrong** | Uses 2 gunicorn workers (the publisher needs exactly 1 process). Installs `fonts-dejavu-core`, which probably lacks the Condensed font the renderer prefers (see `docs/04-rendering-design.md`). Never built. |

Run it only to look around:

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
python seed_demo.py        # optional sample data
python app.py              # http://localhost:8080
```
