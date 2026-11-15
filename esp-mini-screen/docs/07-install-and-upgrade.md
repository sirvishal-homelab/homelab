# 07 - Install and upgrade runbook (server, Docker)

Status: **DRAFT, updated 2026-11-15 for S-1 (code exists, CI builds and runs the image). Not yet run on the server.** Describes the agreed path: build on
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
| Docker volume `esp-screen-data` | SQLite database and its backups (`/data` in the container) |
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

**What to expect after S-1:** the screen shows a **"Screen connected - waiting for Beszel"** card with
the start time, within a few seconds of `up -d`. The only URL is `/healthz`; the settings page arrives
in S-4 and live status in S-2. If the log warns that album autoplay is on, turn off **Enable image
auto display** on the screen's Pictures page.

Later (S-4): open `http://<server-ip>:8081/` from your phone, pick the focus systems, Save.

## Backups

The database only holds settings and container memory; losing it means re-picking focus
systems. The app keeps 7 nightly copies in the volume. To keep one off the server:

```bash
mkdir -p /root/esp-screen-backups
docker run --rm -v esp-screen-data:/data:ro -v /root/esp-screen-backups:/out alpine \
    sh -c 'cp -n /data/backups/*.db /out/'
```

Restore drill (S-5): stop the app, copy a backup over `/data/screen.db`, start, check the
Settings page shows the same focus systems.

## Upgrade

```bash
cd /root/homelab/esp-mini-screen/app/deploy

# 1. Note what runs now (for rollback)
git -C /root/homelab rev-parse --short HEAD

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

```bash
cd /root/homelab
git checkout <previous-short-sha> -- esp-mini-screen
cd esp-mini-screen/app/deploy
docker compose build && docker compose up -d
```

If the newer version applied a migration, first restore its `pre-migration` backup (as in the
restore drill). Return to `main` with `git -C /root/homelab checkout main -- esp-mini-screen`
once a fix is merged.

## Screen firmware

Do **not** update the screen's firmware casually: the app depends on its undocumented upload API
(`03`, tested on Ultra-V9.0.50). If you update it, check afterwards that Push now on the
Settings page still works, and record the new version in `03`.

## Monitoring

Beszel already watches this container host. Add (owner, S-5): container `esp-screen` health
shown as healthy; alert if it is not, where your Beszel version supports container alerts.

## Uninstall

```bash
cd /root/homelab/esp-mini-screen/app/deploy
docker compose down          # add -v to also delete settings and backups
```

Then switch the screen back to a weather theme on its own web page, and delete the Beszel app user.
