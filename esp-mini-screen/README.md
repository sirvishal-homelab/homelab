# ESP Mini Screen

What to do with a **GeekMagic "SmallTV"** desktop display (a 1.54" 240x240 colour IPS
screen driven by an ESP8266 or ESP32, sold as a "WiFi weather clock") so that it is useful
in this homelab and home.

> **Status: running.** Homelab status screen fed by Beszel on a SmallTV-Ultra (stock firmware
> Ultra-V9.0.50, nothing flashed). App 0.5.0 on the Terminus container: live overview, alert
> cards, settings page, update dots. Next: S-5 (operate and verify, `docs/06`).

## Repository layout

```
esp-mini-screen/
├── README.md                  This file
├── CLAUDE.md                  Working rules for Claude Code in this folder
└── docs/
    ├── 00-overview.md         Device, goals, constraints, what the lab already has
    ├── 01-ideas.md            Candidate uses, scored against the existing homelab
    ├── 02-open-questions.md   What must be answered before choosing a direction
    ├── 03-device-api.md       Stock firmware HTTP API (spike results)
    ├── 04-requirements.md     Homelab status screen fed by Beszel (agreed)
    ├── 05-design.md           Architecture, rules, screens, data model (signed off)
    ├── 06-implementation-plan.md  Gate checklist and build slices S-0..S-5 (signed off)
    ├── 07-install-and-upgrade.md  Install, upgrade, rollback on the Terminus container (partly verified)
    ├── 08-beszel-api.md       Beszel 0.21.2 fields and findings (S-0)
    ├── assets/                Mock-ups (sample data) and their generator
    └── decisions/             Architecture Decision Records (0001-0004 accepted)
app/                           The status-screen service (Python, Docker); see app/README.md
```

Gate clear 2026-11-15. S-0 to S-4 done 2026-11-15; S-4b (update dots) done 2026-11-21.

## Related

- [`../trmnl-calendar/`](../trmnl-calendar/): the e-paper calendar. Its Pillow rendering
  and push approach (ADR-0002, ADR-0006) is the likely template for this project.
- [`../trmnl-calendar/docs/13-infrastructure.md`](../trmnl-calendar/docs/13-infrastructure.md):
  the current infrastructure inventory.
