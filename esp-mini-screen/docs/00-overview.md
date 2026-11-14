# 00 - Overview

## Purpose

Find a use for a small always-on colour display that adds real value at home or in the lab,
and decide whether that use is worth building.

## The device

| Item | Value | Confidence |
|------|-------|------------|
| Product | GeekMagic "SmallTV" mini WiFi weather clock (AliExpress item 1005006728045097) | Verified (purchase) |
| Model | **Ultra** (ESP8266 / ESP-12E) or **Pro** (ESP32, touch button) | Pending: check the box or the device web page |
| Screen | 1.54" IPS, 240x240, ST7789 controller, colour | Inferred (both models) |
| Power | USB, always on | Inferred |
| Stock firmware | Weather, clock, photo album / GIF modes; web UI at `http://<screen-ip>/`; accepts image uploads over HTTP; has a firmware upload page | Inferred from community reports; to verify |
| Custom firmware | ESPHome works on both. On the Ultra (ESP8266) RAM is tight and needs a non-buffered display driver; the Pro (ESP32) can run LVGL | Inferred from community reports |

## What the lab already has (from `../trmnl-calendar/docs/13-infrastructure.md`)

- 2-node **Proxmox** cluster (one node is a Dell OptiPlex 9020).
- **Pi-hole**: DNS and DHCP for the LAN.
- **Beszel** hub with agents: host and container monitoring.
- Docker in LXC containers; **Watchtower** for image updates.
- **Terminus** + **TRMNL OG** e-paper display, and the **TRMNL Calendar** app (Python, FastAPI,
  Pillow rendering, pushes images on a schedule).

Not present: **Home Assistant** (so HA integrations such as `geekmagic-hacs` and ESPHome's HA API
only apply if HA is added).

## Goals (draft, confirm with owner)

1. Show something worth glancing at, without opening a phone or a dashboard.
2. Low maintenance: no new fragile dependency; survives reboots and Wi-Fi drops.
3. Fit the existing patterns (Docker service, config in git, secrets in env, Beszel for health).

## Constraints

- 240x240 is tiny: one to four values per screen, large type, colour used for status.
- The display is colour and refreshes fast, which complements the slow greyscale TRMNL.
- The repo is public: no addresses or secrets in docs.
