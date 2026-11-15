# ESP Mini Screen

What to do with a **GeekMagic "SmallTV"** desktop display (a 1.54" 240x240 colour IPS
screen driven by an ESP8266 or ESP32, sold as a "WiFi weather clock") so that it is useful
in this homelab and home.

> **Status: exploring.** Device confirmed as SmallTV-Ultra (firmware Ultra-V9.0.50).
> No direction chosen yet. Nothing has been flashed or built.

## Repository layout

```
esp-mini-screen/
├── README.md                  This file
├── CLAUDE.md                  Working rules for Claude Code in this folder
└── docs/
    ├── 00-overview.md         Device, goals, constraints, what the lab already has
    ├── 01-ideas.md            Candidate uses, scored against the existing homelab
    ├── 02-open-questions.md   What must be answered before choosing a direction
    └── decisions/             Architecture Decision Records (none yet)
```

Planned once a direction is chosen: `docs/03-requirements.md`, `docs/04-design.md`, then
code in `app/` (if a push service is built) or `firmware/` (if the device is reflashed).

## Related

- [`../trmnl-calendar/`](../trmnl-calendar/): the e-paper calendar. Its Pillow rendering
  and push approach (ADR-0002, ADR-0006) is the likely template for this project.
- [`../trmnl-calendar/docs/13-infrastructure.md`](../trmnl-calendar/docs/13-infrastructure.md):
  the current infrastructure inventory.
