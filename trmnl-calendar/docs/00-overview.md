# 00 - Overview

## Purpose

Show a personal calendar on a TRMNL e-paper display using only hardware and software on a
home network. No cloud account is required for the calendar content.

## What the user asked for (in order)

1. Run a **local Terminus server** so the TRMNL device is served from the home network.
2. Build a **calendar project** that updates on a schedule and is shown on the device.
3. A **web front end** to enter events, stored in a **database**, and rendered to an **image**
   the device can display.
4. Terminus must get the image from the app (see `05-terminus-integration.md`).
5. The **image is customisable**: choose a **week, two-week or month** view, and divide each
   calendar cell into **sections** such as "events", "work", "lunch", "dinner".
6. All configuration is **modular** and **stored in the database** so changes are not lost.
7. Pause coding; carry the design into a project and settle **architecture and design first**.

## Scope

In scope: event entry and storage, recurrence, sections, layout configuration, image
rendering, delivery to Terminus, a home-network deployment.

Out of scope (for now, confirm with the user): multi-user accounts, internet exposure,
importing from Google Calendar or ICS feeds, notifications, mobile app, TRMNL's hosted cloud
service as a target.

## Glossary

| Term | Meaning |
|------|---------|
| TRMNL | The e-paper display device (and the company behind it). |
| Terminus | The "flagship" self-hosted server for TRMNL devices (BYOS: bring your own server). Ruby/Hanami, PostgreSQL, Valkey, Sidekiq. |
| BYOS | Bring Your Own Server. A server that speaks the device protocol in place of TRMNL's cloud. |
| Screen | In Terminus, one image (or HTML page) that can be shown on a device. |
| Playlist | In Terminus, the ordered set of screens a device cycles through. |
| Extension | In Terminus, a scheduled plugin that fetches data and renders a Liquid template into a screen. |
| Section | A named row inside each day cell (for example "Lunch"). Has a relative height and can be hidden. |
| View / layout | How the image is composed: week, two-week, month, agenda, or month-plus-agenda. |
| Push | The app telling Terminus to fetch a fresh image by creating or updating a screen via its API. |

## Constraints that shape everything

- The display is 1-bit black and white (the original device is 800x480). Anything drawn must
  survive reduction to pure black and white. See `04-rendering-design.md`.
- Terminus is beta. Its API may change between versions.
- The app runs on a home network, probably on a small always-on machine.
- The user wants the design settled before more code is written.
