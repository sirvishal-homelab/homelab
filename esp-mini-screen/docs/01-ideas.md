# 01 - Ideas

Status: **Draft for discussion.** Effort: S = an evening, M = a weekend, L = several weekends.

## Two ways to drive the screen

| Approach | How | Fits this lab? |
|----------|-----|----------------|
| **A. Push images to stock firmware** | A small service renders a 240x240 PNG/JPEG with Pillow and uploads it to the device over HTTP. Same pattern as TRMNL Calendar (app owns data and drawing; the device only displays). | **Yes.** No flashing, reuses existing code style, Docker, CI. Works on Ultra and Pro. |
| **B. Reflash with ESPHome** | Device becomes programmable; draws its own UI; Pro can use LVGL and the touch button. | Best with Home Assistant, which the lab does not run. One-way (stock firmware may need an activation code to restore). |

Recommendation: start with **A**. Revisit B only if Home Assistant is added or touch input
becomes important.

## Candidate uses

| # | Idea | Data sources (already in the lab) | Value | Effort |
|---|------|-----------------------------------|-------|--------|
| 1 | **Homelab status light**: green "all good" most of the time; red with the failing item when not. Hosts up, CPU/RAM/disk, container down. | Beszel hub API (PocketBase), Proxmox API | High: you notice problems without checking | M |
| 2 | **Calendar companion**: "Next up" event and today's count, from TRMNL Calendar. Colour and instant, where the TRMNL is greyscale and 15 min behind. | TRMNL Calendar API (add a small read endpoint if needed) | Medium-high, household-facing | S-M |
| 3 | **Pi-hole panel**: queries today, % blocked, status; red if Pi-hole is down (which also means DNS and DHCP are down). | Pi-hole API | Medium; Pi-hole is a single point of failure | S |
| 4 | **TRMNL health**: TRMNL battery estimate and last check-in; warn before it goes flat or stops updating. | TRMNL Calendar `/api/display/device` (battery module already exists) | Medium | S |
| 5 | **Rotating carousel** of 1-4 with an alert override (any red item jumps to the front). | All of the above | Highest overall | M (after 1-4) |
| 6 | Weather + clock (stock firmware, zero work) | Built-in | Low; already what it does | none |

## Suggested direction (to confirm)

A single **`esp-mini-screen` push service** (Docker, Python, Pillow), deployed next to TRMNL
Calendar. Phase 1: idea 1 (status light) with idea 3 folded in. Phase 2: add 2 and 4 and the
carousel with alert override.

## Before choosing

See `02-open-questions.md`. The model (Ultra or Pro) and the stock firmware's upload API
must be verified on the real device first; that is a throwaway spike, not production code.
