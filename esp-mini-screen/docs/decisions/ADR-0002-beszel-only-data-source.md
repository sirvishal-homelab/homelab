# ADR-0002: Beszel is the only data source
Status: **Accepted 2026-11-15 by the owner**
Date: 2026-11-14

## Context
Beszel agents run on both Proxmox nodes, TrueNAS and other hosts. Beszel reports systems and
Docker containers but not Proxmox VMs or LXC guests. The owner wants 3 focus systems plus a
container count, and will create a dedicated read-only Beszel account.

## Decision
Read only from the Beszel hub's PocketBase REST API as a read-only user. No Proxmox API.

## Consequences (good and bad)
- Good: one credential, one integration, read-only; matches what the owner already watches.
- Bad: a stopped VM or LXC without an agent is invisible.
- Bad: Beszel's API may change in minor releases; mitigated by fixtures and one mapping module.

## Alternatives considered
- Add the Proxmox API (PVEAuditor token) for per-guest state: deferred; can be added later.
- Prometheus/Grafana: not in the lab; heavier.
