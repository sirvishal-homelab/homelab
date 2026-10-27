# Architecture Decision Records

One file per decision. Status values: **Proposed**, **Accepted**, **Superseded**.
Statuses are as of 2026-10-27. All ten ADRs are accepted (owner design sign-off 2026-10-27). ADR-0003 keeps review gate E-1;
ADR-0008 is accepted for stage 1 only.

| ADR | Title | Status |
|-----|-------|--------|
| 0001 | Design first; the prototype is reference only | **Accepted** |
| 0002 | Separate calendar app that owns data and drawing; Terminus is delivery | **Accepted** |
| 0003 | Push the image to Terminus through its Screens API | **Accepted**, review gate E-1 |
| 0004 | Sections (shown as categories): weighted lanes or banners, in the database | **Accepted** |
| 0005 | Configuration lives in the database; secrets stay in the environment | **Accepted** |
| 0006 | Render with Pillow in two phases to a grayscale 1- or 2-bit PNG | **Accepted** |
| 0007 | SQLite, one process and one worker, push state in the database | **Accepted** |
| 0008 | Delivery pipeline: CI first, manual pull deploys | **Accepted** (stage 1) |
| 0009 | Python 3.12, FastAPI, server-rendered pages, no build step | **Accepted** |
| 0010 | Floating local time, clamped monthly repeats, skip exceptions | **Accepted** |

Template:

```
# ADR-NNNN: Title
Status: Proposed | Accepted | Superseded by ADR-XXXX
Date:
## Context
## Decision
## Consequences (good and bad)
## Alternatives considered
```
