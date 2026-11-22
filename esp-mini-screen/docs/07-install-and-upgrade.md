# 07 - Install and upgrade runbook (server, Docker)

Status: **Partly verified 2026-11-15:** install and upgrades 0.1.0 → 0.4.0 ran on the server as
written. Backup, restore drill and rollback corrected 2026-11-15 (volume name, WAL files,
image-tag rollback); to be run in S-5. Describes the agreed path: build on
the server from a clone, manual pull and update (same as TRMNL Calendar `docs/19`). Slice S-1
runs every command on the real server and corrects this file; S-5 marks it Verified.

Host: the **existing Terminus container** (Ubuntu 24.04 LXC with Docker), next to TRMNL
Calendar. Placeholders: `<server-ip>`, `<screen-ip>`, `<beszel-url>`. Real addresses stay out of
this public repo.

## Layout on the server

| Path | What |
|------|------|
| `/root/terminus` | Terminus (unchanged) |
| `/root/homelab` | Existing sparse clone of this repo, already used by TRMNL Calendar |
| `/root/homelab/esp-mini-screen/app/deploy/.env` | Secrets and settings, `chmod 600`, never committed |
| Docker volume `esp-screen_esp-screen-data` | SQLite database and its backups (`/data` in the container) |
| Port `8081` | Settings page (TRMNL Calendar keeps `8080`, Terminus `2300`) |

Expected footprint: one container, roughly 60-100 MB RAM, an image of about 200 MB.

## Before install

1. **Room on the container.** `df -h /` (want >= 3 GB free), `free -m` (want >= 200 MB available),
   `ss -tln | grep 8081` (want nothing).
2. **Beszel app user** (in the Beszel hub, **menu names vary by version; to verify in S-0**):
   1. Create a user, e.g. `esp-screen@<your-domain>`, with a long random password.
   2. Set its role to **read-only**.
   3. Share every system with it. **Done 2026-11-15** with `SHARE_ALL_SYSTEMS=true` on the hub,
      which shares all systems (including future ones) with every user. The hub runs as the
      systemd service `beszel-hub` (binary in `/opt/beszel`, port 8090), so the variable is a
      drop-in override:
      ```bash
      cat > /etc/systemd/system/beszel-hub.service.d/override.conf <<'EOF'
      [Service]
      Environment="SHARE_ALL_SYSTEMS=true"
      EOF
      systemctl daemon-reload && systemctl restart beszel-hub
      systemctl show beszel-hub -p Environment      # expect SHARE_ALL_SYSTEMS=true
      ```
      The `[Service]` line is required; without it systemd logs "Assignment outside of section.
      Ignoring." Alternative: share system by system in the PocketBase admin (`/_/`, collection
      `systems`, field `users`); new systems then need sharing by hand.
   4. Check: log in as that user in a private window; you should see the systems and be unable to add or edit.
3. **Screen.**
   1. Pi-hole DHCP reservation for the screen (keeps `<screen-ip>` fixed).
   2. Good power: 5 V / 1 A or better adapter.
   3. On `http://<screen-ip>/`: **Settings** → theme **Photo Album**; turn **Auto Switch Themes** off.
      **Pictures** → **Enable image auto display** off, SAVE.
   4. Optional: delete the sample images and `claude-test.jpg` on the Pictures page (frees flash).
4. From the container, check it can reach both: `curl -sI http://<screen-ip>/ | head -1` and
   `curl -s <beszel-url>/api/health`.

## First-time install

```bash
# 1. Add this project to the existing sparse clone
cd /root/homelab
git pull --ff-only
git sparse-checkout add esp-mini-screen
cd esp-mini-screen/app/deploy

# 2. Settings file
cp ../.env.example .env
chmod 600 .env
nano .env     # BESZEL_URL, BESZEL_USER, BESZEL_PASSWORD, SCREEN_URL, HOST_PORT=8081, TZ=America/Halifax

# 3. Build and start
docker compose build
docker compose up -d

# 4. Check
docker compose ps                         # "healthy" after ~30 s
curl -s http://localhost:8081/healthz     # "status":"ok"; "screen": {"last_result": "ok", ...}
docker compose logs --tail=50 esp-screen  # migrations applied, "start-up frame shown on the screen"
```

**After S-4b (0.5.0):** no `.env` change. A small dot appears after a focus system's name when
Beszel reports pending OS updates for it: red for security updates, amber for others. To check:
update one of those systems (`apt update && apt full-upgrade`); the agent re-checks hourly, so the
dot clears within an hour. Turn the dots off on the settings page (Tuning) if you don't want them.

**After S-4 (0.4.0):** open `http://<server-ip>:8081/` on your phone. Your current focus systems
are taken over from `FOCUS_SYSTEMS` automatically; from now on change them on the page (row 1-4,
display name, ignore alerts), plus containers that may stop, timing and thresholds. Save applies
within seconds. You can then delete the `FOCUS_SYSTEMS` line from `.env` (it is ignored).

**After S-3 (0.3.0):** no `.env` change. Problems now show on the screen: a red **ATTENTION**
card alternates with the overview every 10 s while a system is down or an expected container is
gone; an amber, then red, **Beszel unreachable** card replaces the overview after two failed
reads; **Beszel data changed** appears if a Beszel update changes its data. `/healthz` gains
`alert` and `beszel.failures_in_a_row`. To try it safely: stop a container you can spare on the
Docker host (`docker stop watchtower`), wait about 2 minutes for the card, then
`docker start watchtower`; the card clears within a minute.

**After S-2 (0.2.0):** add `FOCUS_SYSTEMS=<name>,<name>,<name>` to `.env` (names as shown in
Beszel, 1 to 4, in order), then `docker compose up -d`. The start-up card shows for about a minute,
then the live overview. `/healthz` gains `beszel` (last read, last error) and `overview` (what the
screen shows). If Beszel cannot be reached the screen keeps its last picture and `/healthz` says
why; the on-screen warning arrives in S-3.

**What to expect after S-1:** the screen shows a **"Screen connected - waiting for Beszel"** card with
the start time, within a few seconds of `up -d`. The only URL is `/healthz`; the settings page arrives
in S-4 and live status in S-2. If the log warns that album autoplay is on, turn off **Enable image
auto display** on the screen's Pictures page.

Later (S-4): open `http://<server-ip>:8081/` from your phone, pick the focus systems, Save.

## Backups

The database only holds settings and container memory; losing it means re-picking focus
systems. The app keeps 7 automatic copies in the volume (`/data/backups`), plus one before every
migration. The volume is called **`esp-screen_esp-screen-data`** (compose project + volume name).

Back up now and copy the backups off the container's volume:

```bash
cd /root/homelab/esp-mini-screen/app/deploy
docker compose exec esp-screen python -m esp_screen.storage.backup --now
mkdir -p /root/esp-screen-backups
docker run --rm -v esp-screen_esp-screen-data:/data:ro -v /root/esp-screen-backups:/out alpine \
    sh -c 'cp -n /data/backups/*.db /out/ && ls -l /out'
```

### Restore (and the S-5 drill)

```bash
cd /root/homelab/esp-mini-screen/app/deploy
docker compose stop esp-screen
# list backups, newest last
docker run --rm -v esp-screen_esp-screen-data:/data alpine ls -1 /data/backups
# restore one: replace the database AND delete its WAL side files, then fix ownership
docker run --rm -v esp-screen_esp-screen-data:/data alpine sh -c '
  cp /data/backups/<backup-file>.db /data/screen.db &&
  rm -f /data/screen.db-wal /data/screen.db-shm &&
  chown 10001:10001 /data/screen.db'
docker compose start esp-screen
curl -s http://localhost:8081/healthz
```

Check on the settings page that the focus systems, names and tuning are the ones from the
backup. **Drill:** change one display name, take a backup, change it again, restore the backup:
the first name must come back.

## Upgrade

```bash
cd /root/homelab/esp-mini-screen/app/deploy

# 1. Keep the running image for a quick rollback
docker tag esp-mini-screen:local esp-mini-screen:previous

# 2. Fetch and read what changed
git -C /root/homelab pull --ff-only
git -C /root/homelab log --oneline -10 -- esp-mini-screen

# 3. If .env.example gained a variable, add it to .env

# 4. Rebuild and restart (migrations run on start, after an automatic pre-migration backup)
docker compose build
docker compose up -d

# 5. Verify
docker compose ps
curl -s http://localhost:8081/healthz
```

Upgrading TRMNL Calendar or Terminus does not touch this app, and the reverse. Both apps share
the clone: `git pull` updates both folders, but each only changes when you rebuild it.

## Roll back

Fast path (no git): go back to the image tagged before the last upgrade.

```bash
cd /root/homelab/esp-mini-screen/app/deploy
docker tag esp-mini-screen:previous esp-mini-screen:local
docker compose up -d --no-build
curl -s http://localhost:8081/healthz      # "version" shows the older number
```

If the newer version applied a database migration (the log says "migrations applied" with a
new number), also restore its `...-pre-migration-v<N>.db` backup as above, because an older app
cannot read a newer schema. To go forward again: `docker compose build && docker compose up -d`.

## Screen firmware

Do **not** update the screen's firmware casually: the app depends on its undocumented upload API
(`03`, tested on Ultra-V9.0.50). If you update it, check afterwards that Push now on the
Settings page still works, and record the new version in `03`.

## Monitoring

- **The screen watches itself:** the header time is the time of the last good Beszel read and
  changes every minute. A time more than a few minutes old means the app or the screen is stuck.
- **Beszel:** the Terminus container host is already monitored; `esp-screen` appears in its
  container list. If your Beszel version offers alerts for a single container, add one for
  `esp-screen`; otherwise the stale header time is the signal.
- **`/healthz`:** `screen.last_result`, `beszel.last_error` and `alert` say what is happening.
- **Memory:** `docker stats --no-stream esp-screen` (expected well under 150 MB).

## Uninstall

```bash
cd /root/homelab/esp-mini-screen/app/deploy
docker compose down          # add -v to also delete settings and backups
```

Then switch the screen back to a weather theme on its own web page, and delete the Beszel app user.
