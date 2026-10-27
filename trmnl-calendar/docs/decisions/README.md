# Architecture Decision Records

One file per decision. Status values: **Proposed**, **Accepted**, **Superseded**.
Statuses are as of 2026-10-27. ADR-0003 and ADR-0008 are accepted; the rest await the owner's sign-off in
phase 2 of the roadmap.

| ADR | Title | Status |
|-----|-------|--------|
| 0001 | Design first; the prototype is reference only | Proposed |
| 0002 | Separate calendar app that owns data and drawing; Terminus is delivery | Proposed |
| 0003 | Push the image to Terminus through its Screens API | **Accepted**, review gate E-1 |
| 0004 | Sections (shown as categories): weighted lanes or banners, in the database | Proposed |
| 0005 | Configuration lives in the database; secrets stay in the environment | Proposed |
| 0006 | Render with Pillow in two phases to a grayscale 1- or 2-bit PNG | Proposed |
| 0007 | SQLite, one process and one worker, push state in the database | Proposed |
| 0008 | Delivery pipeline: CI first, manual pull deploys | **Accepted** (stage 1) |
| 0009 | Python 3.12, FastAPI, server-rendered pages, no build step | Proposed |
| 0010 | Floating local time, clamped monthly repeats, skip exceptions | Proposed |

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
