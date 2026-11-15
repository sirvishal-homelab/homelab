# ADR-0001: Push rendered images to the stock firmware
Status: **Accepted 2026-11-15 by the owner**
Date: 2026-11-14

## Context
The device is a SmallTV-Ultra (ESP8266, no touch, firmware Ultra-V9.0.50). The lab has no
Home Assistant. The stock firmware accepts JPEG uploads and can be told which image to show
(`03-device-api.md`, verified on the device).

## Decision
A service on the lab renders 240x240 JPEGs with Pillow and pushes them over HTTP
(`/doUpload`, `/set?img=`). The screen stays on stock firmware, Photo Album theme.

## Consequences (good and bad)
- Good: no flashing; reversible; same pattern and stack as TRMNL Calendar; any content is possible.
- Good: the screen keeps working as a weather clock if the service is removed.
- Bad: depends on an undocumented API that a firmware update could change.
- Bad: no authentication on the device; the service must check the theme stays correct.

## Alternatives considered
- ESPHome firmware: RAM-limited on ESP8266, best with Home Assistant (absent), one-way.
- geekmagic-hacs integration: requires Home Assistant.
