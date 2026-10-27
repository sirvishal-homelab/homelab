# ADR-0008: Delivery pipeline - CI first, manual pull deploys
Status: **Accepted 2026-10-25** (stage 1); later stages proposed
Date: 2026-10-25

## Context
`homelab` is a public repo with no existing deployments. The server is x86_64, runs Docker and
already runs Watchtower. The owner is new to CI/CD and is happy to update by hand.

## Decision
Stage 1 now: GitHub Actions runs lint, tests, image snapshots, fake-Terminus contract tests and
a Docker build on every pull request; `main` is protected. The owner deploys with
`docker compose pull && docker compose up -d`. Later: publish an amd64 image to GHCR (stage 2),
opt the container into Watchtower (stage 3), add Beszel alerts and backups (stage 4). Details in
`docs/12-delivery-pipeline.md`.

## Consequences
Good: lowest risk, no inbound exposure, no self-hosted runner, fast iteration.
Bad: manual step per release until stage 3; no automatic rollback (pin the previous tag).

## Alternatives considered
Self-hosted runner (unsafe for a public repo); SSH push from Actions (needs inbound path);
a custom deploy timer (not needed with Watchtower present).
