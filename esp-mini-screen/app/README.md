# esp-mini-screen app

Renders a 240x240 homelab status image from Beszel and pushes it to a GeekMagic SmallTV-Ultra.
Design: `../docs/05-design.md`. Install and upgrade: `../docs/07-install-and-upgrade.md`.

**Slice S-1 (this version):** start-up, database and backups, `/healthz`, and a "Screen
connected" frame pushed to the screen at start. Live Beszel status arrives in S-2.

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
