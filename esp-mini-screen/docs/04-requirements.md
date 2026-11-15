# 04 - Requirements (draft)

Status: **Draft, 2026-11-14. Not signed off.** Direction: approach A (push images to stock
firmware), first use is a **homelab status screen** fed by Beszel.

## Context from the owner

- Location: office, near the desk or the server rack (may move). Viewer: the owner only.
- Wants a hook into **Beszel** showing what is up (VMs, containers) and some stats.
- Will create a **dedicated Beszel account** for the app.

## What the data sources can give

| Source | Gives | Does not give |
|--------|-------|---------------|
| **Beszel hub** (PocketBase REST API, user login) | Each system with an agent: status up/down/paused, CPU, memory, disk, uptime. Docker/Podman containers per system with CPU and memory | Proxmox VMs or LXC guests (Beszel has no Proxmox API integration; guests only appear if an agent runs inside them) |
| **Proxmox VE API** (read-only API token, `PVEAuditor` role) | Every VM and LXC across both nodes with running/stopped state; node and cluster quorum status | Docker containers inside guests |

So "VMs/LXCs up" needs the Proxmox API; "hosts and Docker containers" comes from Beszel.
Proxmox is proposed as **optional, phase 1b**: Beszel alone is enough to start.

Beszel note: its docs warn the API data structure can change in minor releases. Pin the hub
version the app is tested against and keep the Beszel client in one small module with tests.

## Functional requirements

| ID | Requirement | Phase |
|----|-------------|-------|
| F-1 | Read all systems from Beszel: name, status, CPU %, memory %, disk %. | 1 |
| F-2 | Read Docker containers from Beszel: count running vs expected, per host. | 1 |
| F-3 | Render a 240x240 **overview** screen: summary tiles (guests up, containers up) and one row per host with status dot, CPU, memory and disk bars. See `assets/mock-overview.jpg`. | 1 |
| F-4 | Colour by threshold: green normal, amber >= 75 %, red >= 90 % (configurable). | 1 |
| F-5 | **Alert override**: if any red condition holds, show a full-screen alert (`assets/mock-alert.jpg`) and alternate with the overview. | 1 |
| F-6 | "Expected" containers: a container that was running and disappears counts as down (the app remembers the last known set; an allow-list can mark ones that are allowed to stop). | 1 |
| F-7 | Push only when the rendered image changes; otherwise re-push at most every 15 minutes as a keep-alive. | 1 |
| F-8 | Before pushing, check the device is on the Photo Album theme with autoplay off; log a warning if not (set it automatically once the request is known). | 1 |
| F-9 | Read VM/LXC state from the Proxmox API and fill the "VMs/LXC" tile; a guest marked "onboot" that is stopped counts as down. | 1b |
| F-10 | If Beszel cannot be reached, show a distinct "Beszel unreachable" screen (grey/amber), not stale green. | 1 |
| F-11 | A minimal status page or `/health` endpoint so Beszel or the owner can see the app itself is alive. | 1 |

Red conditions (draft, owner to confirm): a host down; a Proxmox node down or quorum lost (1b);
an expected container or onboot guest down; disk >= 90 %; Beszel unreachable for > 5 min.

## Non-functional

| ID | Requirement |
|----|-------------|
| N-1 | Runs as one small Docker container on the existing Docker LXC, next to TRMNL Calendar. |
| N-2 | Python 3.12 + Pillow (same stack and conventions as TRMNL Calendar). |
| N-3 | Config (thresholds, poll interval, device address) in a config file or env; **secrets only in env** (`BESZEL_URL`, `BESZEL_USER`, `BESZEL_PASSWORD`, later `PVE_TOKEN_ID`, `PVE_TOKEN_SECRET`). Nothing secret in git. |
| N-4 | Poll every 60 s (configurable). Upload under 2 s. A change appears on screen within ~1 min. |
| N-5 | Recovers by itself after the screen, Beszel or the network restarts. |
| N-6 | Read-only everywhere: the Beszel account and Proxmox token cannot change anything. |
| N-7 | Unit tests for rendering and the Beszel and Proxmox clients (with recorded fixtures); CI like TRMNL Calendar. |

## Owner set-up (when we build)

1. Beszel: create user (e.g. `esp-screen`), role **read-only**, and share all systems with it.
2. (1b) Proxmox: create user `monitoring@pve` and an API token with the `PVEAuditor` role.
3. Pi-hole: DHCP reservation for the screen (Q-8).
4. Screen: Photo Album theme, album autoplay off, theme auto-switch off; delete the sample images.

## Out of scope (for now)

Calendar or TRMNL-health screens (ideas 2 and 4), Pi-hole stats (idea 3), night mode, touch or
buttons, Home Assistant.
