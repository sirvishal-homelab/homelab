# 08 - Beszel API (S-0 contract spike)

Status: **Recorded 2026-11-15 against Beszel 0.21.2** (hub and all agents), as the read-only
app user (`role: readonly`). Read-only calls only. Addresses removed; system and container
names replaced with generic ones (owner decision). Fixtures: `app/tests/fixtures/beszel-0.21.2/`.

## Access

- Hub: `http://beszel:8090` (resolves via Pi-hole from the Terminus container).
- Auth: `POST /api/collections/users/auth-with-password` → token; send as `Authorization: <token>`.
- All systems visible to the app user via `SHARE_ALL_SYSTEMS=true` (the `users` field on each
  system still lists only the admin; sharing is applied by the hub, not stored per record).

## `GET /api/collections/systems/records?perPage=200`

| Field | Meaning | Example |
|-------|---------|---------|
| `id` | Stable record id (use as key) | `as4fkbiyv9cpbkx` |
| `name` | Display name | `node-a` |
| `status` | `up`, `down`, `paused`, `pending` | `up` |
| `updated` | Last update (UTC) | `2026-11-15 10:23:33.567Z` |
| `info.cpu` | CPU % | `0.91` |
| `info.mp` | Memory % | `6.97` |
| `info.dp` | Disk % of the agent's **root** filesystem | `53.91` |
| `info.u` | Uptime, seconds | `2424250` |
| `info.v` | Agent version | `0.21.2` |
| `info.sv` | Systemd services `[total, failed]` (not on every system) | `[39, 1]` |
| `info.la` | Load average 1/5/15 | `[0.04, 0.02, 0.02]` |
| `info.dt` | Highest temperature (°C), not used | `31` |
| `info.bb`, `ct`, `t`, `pu` | Bandwidth, cores/threads and other values, not used | |

A paused system keeps zeroed `info` and an old `updated`.

## `GET /api/collections/containers/records?perPage=500`

| Field | Meaning | Example |
|-------|---------|---------|
| `id` | Docker short id: **changes when a container is recreated**, do not use as key | `cae8e105d248` |
| `system` | System record id | `grh0kefzwq2t2g3` |
| `name` | Container name (key = `system` + `name`) | `calendar` |
| `status` | Docker status text | `Up 4 days` |
| `health` | Number; `2` on every container seen (meaning not yet confirmed) | `2` |
| `cpu`, `memory` | CPU %, memory MB | `0.54`, `27.54` |
| `image`, `ports`, `updatable`, `updated` | Not used in v1 | |

**Corrected 2026-11-15 (live test):** a stopped container is **not** removed at once. Its record
stays, still saying e.g. `Up 5 days`, but its `updated` time stops moving. Every running
container's `updated` (milliseconds) equals its system's `updated` to the millisecond, because
both are written by the same agent report. Rule used by the app: a container is running only if
its `updated` is within 15 s of its system's `updated`. Verified on the live hub: running
containers 0 s behind, a stopped `watchtower` 480 s behind, and the old record of a recreated
container 540 s behind. Beszel deletes old records later; that case is handled too.

## Detail collections (not needed for v1)

`system_stats` and `container_stats` hold time series (`type` `1m`, ...). Per-system stats include
`d` (disk GB), `du` (used GB), `dp`, and `efs` (extra filesystems) when the agent is configured
with them.

## Findings that change the design

| ID | Finding | Effect |
|----|---------|--------|
| B-1 | Container ids change on recreate | Key containers by system id + name (as designed). |
| B-2 | Stopped containers keep a stale record (corrected 2026-11-15; first read as "vanish") | Count only records updated in the system's latest report; then the missing-for-2-reads rule applies (`beszel/client.py` `running_containers`). |
| B-3 | `nas` (TrueNAS) reports disk 0.05 % of 833 GB: the agent only sees its own small dataset, no `efs` | **Resolved 2026-11-15:** disk is not shown or monitored at all (owner). No agent change needed. |
| B-4 | `backup` is `paused` since 2026-05-31 | **Resolved 2026-11-15:** paused = monitored-off on purpose; excluded from counts and alerts, shown as "n paused". |
| B-5 | `info.sv` gives failed systemd services (`vm-1` has 1) | Possible later red/amber condition; not in v1. |
