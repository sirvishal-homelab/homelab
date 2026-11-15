# 04 - Requirements (draft)

Status: **Complete, agreed with owner 2026-11-15** (direction, layout, red conditions, timing, gaps G-1..G-5 closed). Direction: approach A (push images to stock
firmware), first use is a **homelab status screen** fed by Beszel.

## Context from the owner

- Location: office, near the desk or the server rack (may move). Viewer: the owner only.
- Wants a hook into **Beszel** showing what is up (VMs, containers) and some stats.
- Will create a **dedicated Beszel account** for the app.
- Beszel agents already run on **both Proxmox nodes and TrueNAS**, so those hosts are in Beszel today.
- Wants a **lightweight web interface** to choose which systems are shown.

## What the data sources can give

| Source | Gives | Does not give |
|--------|-------|---------------|
| **Beszel hub** (PocketBase REST API, user login) | Each system with an agent: status up/down/paused, CPU, memory, disk, uptime. Docker/Podman containers per system with CPU and memory | Proxmox VMs or LXC guests (Beszel has no Proxmox API integration; guests only appear if an agent runs inside them) |

So "VMs/LXCs up" needs the Proxmox API; "hosts and Docker containers" comes from Beszel.
**Decision (owner, 2026-11-14): Beszel only.** Agents already run on both Proxmox nodes and TrueNAS, so the infrastructure hosts are covered. Per-VM/LXC state (Proxmox API) is out of scope; the screen counts containers instead.

Beszel note: its docs warn the API data structure can change in minor releases. Pin the hub
version the app is tested against and keep the Beszel client in one small module with tests.

## Functional requirements

| ID | Requirement | Phase |
|----|-------------|-------|
| F-1 | Read all systems from Beszel: name, status, CPU %, memory %, disk %. | 1 |
| F-2 | Read Docker containers from Beszel: count running vs expected, per host. | 1 |
| F-3 | Render a 240x240 **overview** screen: two summary tiles, **Containers up / expected** and **Other systems up / total** (systems that are not focus rows; paused systems excluded from both numbers and shown as "n paused"), and rows for **1 to 4 focus systems** chosen on the settings page (core infrastructure, e.g. both Proxmox nodes and the NAS); up to 3 use the roomy layout, 4 switch to a compact layout with both counters in one strip (`assets/mock-overview-4.jpg`) with status dot and **CPU and memory bars only** (no disk; owner decision 2026-11-15). See `assets/mock-overview.jpg`. | 1 |
| F-4 | Bar colours by threshold: green normal, amber >= 75 %, red >= 90 % (configurable). Colour only; CPU and memory never raise an alert. | 1 |
| F-5 | **Alert override**: if any red condition holds, show a full-screen alert (`assets/mock-alert.jpg`) and alternate with the overview. | 1 |
| F-6 | "Expected" containers: a container that was running and disappears counts as down (the app remembers the last known set; an allow-list can mark ones that are allowed to stop). | 1 |
| F-7 | Upload only when a rendered frame changes; the header shows the **"updated at" time** (`HH:MM`, **24-hour**, owner 2026-11-15), so a frame changes at least once a minute. While alerting, the overview and alert frames are both on the screen and the app switches between them with `/set?img=` (no re-upload). | 1 |
| F-8 | Before pushing, check the device is on the Photo Album theme with autoplay off; log a warning if not (set it automatically once the request is known). | 1 |
| F-10 | If Beszel cannot be reached, show a distinct "Beszel unreachable" screen (grey/amber), not stale green. | 1 |
| F-11 | A minimal status page or `/health` endpoint so Beszel or the owner can see the app itself is alive. A frozen screen is recognisable by an old "updated at" time; Beszel's container monitoring covers the app stopping. | 1 |
| F-16 | If Beszel's data does not match the expected shape (missing or renamed fields after a Beszel update), show a distinct **"Beszel data changed"** frame and name the problem on the settings page, instead of showing wrong numbers. | 1 |
| F-17 | **All tuning lives on the settings page** with these defaults: Beszel read interval 60 s; alert/overview flip 10 s; Beszel-unreachable grace 5 min; missed reads before a container is gone 2; bar amber 75 %, red 90 %; focus systems, order, display names, ignore alerts; containers allowed to stop / forget; screen address. Only the Beszel login and bootstrap values are in `.env`. | 1 |
| F-12 | **Settings page** (phone-friendly, LAN only): lists every system Beszel reports; pick **1 to 4 focus systems**, their order, and an optional short display name (e.g. `truenas` → `NAS`). New systems appear unticked. | 1 |
| F-13 | Settings page also edits: amber/red thresholds, poll interval, which containers may stop without alerting (F-6), and the screen address. | 1 |
| F-14 | Settings page shows a **live preview** of the current 240x240 image and a "Push now" button. | 1 |
| F-15 | Non-focus systems are not drawn as rows but are counted in the Other systems tile and still raise alerts, unless "ignore alerts" is ticked for them. **A paused system is never down**: it is monitored-off on purpose, so it is left out of counts and alerts (owner decision 2026-11-15). | 1 |

Red conditions (**confirmed by owner 2026-11-14, revised 2026-11-15**): a system down (paused never counts);
an expected container down; Beszel unreachable for > 5 min. Disk is **not** monitored (owner decision 2026-11-15).

## Non-functional

| ID | Requirement |
|----|-------------|
| N-1 | Runs as one small Docker container on the existing Docker LXC, next to TRMNL Calendar. |
| N-2 | Python 3.12 + Pillow (same stack and conventions as TRMNL Calendar). |
| N-3 | Settings (selected systems, order, names, thresholds, intervals, device address) stored in **SQLite**, edited through the settings page (as TRMNL Calendar ADR-0005/0007); **secrets only in env** (`BESZEL_URL`, `BESZEL_USER`, `BESZEL_PASSWORD`). Nothing secret in git. |
| N-4 | Read Beszel every 60 s (configurable). Upload under 2 s. A change appears on screen within ~1 min. |
| N-10 | Times shown on the screen and settings page in 24-hour format, local time (`TZ=America/Halifax`). |
| N-5 | Recovers by itself after the screen, Beszel or the network restarts. |
| N-6 | Read-only: the Beszel account cannot change anything. |
| N-9 | Install, upgrade, rollback and uninstall steps documented before coding (`07`) and verified on the server before the project is called done. Hosted on the existing Terminus container. |
| N-8 | Settings page has no login (LAN only, single user, same as TRMNL Calendar); it holds no secrets and cannot change Beszel. |
| N-7 | Unit tests for rendering and the Beszel client (with recorded fixtures); CI like TRMNL Calendar. |

## Owner set-up (when we build)

1. Beszel: create user (e.g. `esp-screen`), role **read-only**, and share all systems with it.
2. Pi-hole: DHCP reservation for the screen (Q-8).
3. Screen: Photo Album theme, album autoplay off, theme auto-switch off; delete the sample images.

## Operating notes (no app work)

- **Planned maintenance:** pause the system in Beszel before working on it; paused never counts
  as down, so the screen stays calm. Resume it afterwards.
- **Night:** use the screen's own **Night Mode** (Settings page on the device: start, end,
  night brightness). It works on the Photo Album theme; the app does not manage brightness.

## Out of scope (for now)

Proxmox API / per-VM and LXC state, calendar or TRMNL-health screens (ideas 2 and 4), Pi-hole stats (idea 3), night mode, touch or
buttons, Home Assistant.
