# 19 - Install and upgrade runbook (server, Docker)

Status: **DRAFT, 2026-10-27.** Slice 0 code exists (`/healthz`, migrations, backups); CI builds the image and
checks `/healthz` and the backup command, but it has not yet been run on the server, so these steps describe the agreed path (ADR-0008 stage 1: build on the server from a clone, manual
pull and update). Slice 0 must run every command here on the real server and correct this
file. It will later be copied to `app/deploy/README.md`.

Placeholders: `<server>` is the Docker host, `<terminus>` is `http://<server>:2300`.
Real addresses stay out of this public repo; keep them in your own notes.

## Layout on the server

| Path | What |
|------|------|
| `/root/terminus` | Existing Terminus clone and compose project (unchanged here) |
| `/root/homelab` | Sparse clone of this repo (only `trmnl-calendar/`) |
| `/root/homelab/trmnl-calendar/app/deploy/.env` | Secrets and settings, `chmod 600`, never committed |
| Docker volume `calendar-data` | SQLite database, backups (`/data` in the container) |

## Before the first install

1. Terminus is healthy and the device shows its screen (done).
2. In Terminus, create a dedicated user for the calendar (do not reuse your admin login).
3. DHCP reservations for the server and the device (P-1), so addresses never change.
4. Disk: 20 GB (P-7, done). Check `df -h /` and `docker system df`.
5. Pick the port (default 8080) and check it is free: `ss -tln | grep 8080`.
6. `PUBLIC_BASE_URL` must use the server's **IP address**, not a name, because Terminus
   fetches the image from it and the device network may not resolve your DNS.

## First-time install

```bash
# 1. Get only the calendar folder (public repo, no credentials needed)
cd /root
git clone --filter=blob:none --sparse https://github.com/sirvishal-homelab/homelab.git
cd homelab
git sparse-checkout set trmnl-calendar
cd trmnl-calendar/app/deploy

# 2. Create the settings file
cp ../.env.example .env
chmod 600 .env
nano .env        # set at least: TERMINUS_URL, TERMINUS_LOGIN, TERMINUS_PASSWORD, PUBLIC_BASE_URL

# 3. Build the image from the clone and start
docker compose build
docker compose up -d

# 4. Check
docker compose ps                       # state "healthy" after about 30 s
curl -s http://localhost:8080/healthz   # expect JSON with "ok"
docker compose logs --tail=50 calendar  # migrations applied, fonts found, one worker
```

Then:

1. Open `http://<server>:8080/` on your phone. Add one test event.
2. Pictures and push arrive in slices 1 and 4; until then `/healthz` is the check.
3. In Terminus, make sure the playlist for the device contains only the calendar screen, and
   that **Display Compatibility** stays on for the model (it removes 2-bit ghosting).
4. Wait for the device's next refresh (it polls every 900 s) or trigger one.

## Upgrade

```bash
cd /root/homelab/trmnl-calendar/app/deploy

# 1. What is running now (write this down for rollback)
git -C /root/homelab rev-parse --short HEAD
docker compose images

# 2. Back up the database (the app also does this before a migration)
docker compose exec calendar python -m trmnl_calendar.storage.backup --now   # writes /data/backups/calendar-*-manual.db
#    Or, without the app: docker run --rm -v calendar-data:/data -v "$PWD":/out alpine \
#        sh -c 'cp /data/calendar.db /out/calendar-$(date +%F-%H%M).db'

# 3. Fetch the new code and read what changed
git -C /root/homelab pull --ff-only
git -C /root/homelab log --oneline -10 -- trmnl-calendar

# 4. Rebuild and restart (migrations run on start)
docker compose build
docker compose up -d

# 5. Verify
docker compose ps
curl -s http://localhost:8080/healthz
docker compose logs --tail=30 calendar
```

Migrations are forward-only. If `.env.example` gained a variable, add it to `.env` before step 4.

## Roll back

```bash
cd /root/homelab
git checkout <previous-short-sha> -- trmnl-calendar     # or: git checkout <tag>
cd trmnl-calendar/app/deploy
docker compose build && docker compose up -d
```

If the new version had applied a migration, restore the pre-migration backup first:

```bash
docker compose down
docker run --rm -v calendar-data:/data -v "$PWD":/in alpine \
    sh -c 'cp /in/<backup-file>.db /data/calendar.db && rm -f /data/calendar.db-wal /data/calendar.db-shm'
docker compose up -d
```

Afterwards return the clone to `main` with `git -C /root/homelab checkout main` once a fix is merged.

## Later stages (ADR-0008)

- **Stage 2:** CI publishes an amd64 image to GHCR on merge. The compose file switches from
  `build:` to `image: ghcr.io/sirvishal-homelab/trmnl-calendar:<tag>`; upgrade becomes
  `docker compose pull && docker compose up -d`; rollback pins the previous tag. The package
  must be public (the repo is public) so the server needs no registry login.
- **Stage 3:** opt-in Watchtower label for the calendar only (runs daily 04:00 UTC).

## Terminus itself

Terminus is not deployed from this repo. Its upgrade policy is tracked in `13` (P-2, P-3, P-4):
keep local changes in a `compose.override.yml`, pin the image by digest, remove its Watchtower
labels, and upgrade deliberately:

```bash
cd /root/terminus
docker compose pull && docker compose up -d
docker compose logs --tail=50 web
```

Check that the model (`og_plus`) and the playlist still exist afterwards. The calendar looks
the model up by name, so it survives a changed id.

## Not verified yet

Every command above for the calendar app (CI built and ran the image; the server has not), the `/render/current.png` URL (slice 1), `sparse-checkout` on the
server's git version, the backup command name, and the rollback path. Slice 0 owns these.
