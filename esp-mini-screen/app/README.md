# esp-mini-screen app

Renders a 240x240 homelab status image from Beszel and pushes it to a GeekMagic SmallTV-Ultra.
Design: `../docs/05-design.md`. Install and upgrade: `../docs/07-install-and-upgrade.md`.

**Version 0.2 (slice S-2):** reads Beszel every 60 s and shows the live overview (counters for
containers and other systems, CPU and memory rows for 1-4 focus systems from `FOCUS_SYSTEMS`),
uploading only when the picture changes. Alerts arrive in S-3, the settings page in S-4.

## Layout

```
src/esp_screen/
  config.py        environment (secrets, bootstrap)
  main.py          app factory, start-up, /healthz
  publish.py       render + push + record (the only glue module)
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
