# ESP Mini Screen

What to do with a **GeekMagic "SmallTV"** desktop display (a 1.54" 240x240 colour IPS
screen driven by an ESP8266 or ESP32, sold as a "WiFi weather clock") so that it is useful
in this homelab and home.

> **Status: exploring.** Device confirmed as SmallTV-Ultra (firmware Ultra-V9.0.50).
> Direction: homelab status screen fed by Beszel (design signed off 2026-11-15; building). Nothing flashed or built.

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
    ├── 05-design.md           Architecture, rules, screens, data model (proposed)
    ├── 06-implementation-plan.md  Gate checklist and build slices S-0..S-5 (proposed)
    ├── 07-install-and-upgrade.md  Install, upgrade, rollback on the Terminus container (draft)
    ├── 08-beszel-api.md       Beszel 0.21.2 fields and findings (S-0)
    ├── assets/                Mock-ups (sample data) and their generator
    └── decisions/             Architecture Decision Records (0001-0003 accepted)
```

Gate clear 2026-11-15. S-0 done 2026-11-15. Next: S-1 (foundations and first push)., then
code in `app/` (if a push service is built) or `firmware/` (if the device is reflashed).

## Related

- [`../trmnl-calendar/`](../trmnl-calendar/): the e-paper calendar. Its Pillow rendering
  and push approach (ADR-0002, ADR-0006) is the likely template for this project.
- [`../trmnl-calendar/docs/13-infrastructure.md`](../trmnl-calendar/docs/13-infrastructure.md):
  the current infrastructure inventory.
