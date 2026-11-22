# esp-mini-screen app

Renders a 240x240 homelab status image from Beszel and pushes it to a GeekMagic SmallTV-Ultra.
Design: `../docs/05-design.md`. Install and upgrade: `../docs/07-install-and-upgrade.md`.

**Version 0.5 (slice S-4b):** live overview from Beszel, alert cards, a settings page at `/`
(focus systems, display names, ignore alerts, containers that may stop, timing, thresholds,
update dots on/off, preview, Push now), and a pending-updates dot after focus names (red:
security, amber: other; ADR-0004). Remaining: S-5 (operate and verify).

## Layout

```
src/esp_screen/
  config.py        environment (secrets, bootstrap)
  main.py          app factory, start-up, /healthz
  publish.py       render + push + record (the only glue module)
  beszel/client.py Beszel hub API; all knowledge of its JSON (docs/08)
  domain/          pure rules: snapshot -> status, alerts
  loop.py          read, evaluate, render, push; alert flip
  web/             settings page
  device/client.py the screen's HTTP API (docs/03)
  render/          pure drawing, bundled Inter fonts (OFL)
  storage/         SQLite, migrations, backups, small repositories
tests/
  fakes/screen.py  fake screen firmware (httpx MockTransport)
  unit/, integration/
```

## Develop

```bash
uv sync
uv run ruff check . && uv run ruff format --check .
uv run pytest -q
```

## Run locally

```bash
DATA_DIR=/tmp/esp SCREEN_URL=http://<screen-ip> uv run uvicorn --factory esp_screen.main:app_factory --port 8081
curl -s localhost:8081/healthz
```

Back up now (inside the container): `python -m esp_screen.storage.backup --now`.
