# Architecture Decision Records

One file per decision. Status values: **Proposed**, **Accepted**, **Superseded**.
All ADRs below are **Proposed**: they record what the first session chose and why, and need
the user's sign-off in phase 2 of the roadmap.

| ADR | Title | Status |
|-----|-------|--------|
| 0001 | Design first; the prototype is reference only | Proposed |
| 0002 | Separate calendar app that owns data and drawing; Terminus is delivery | Proposed |
| 0003 | Push the image to Terminus through its Screens API | Proposed (unverified) |
| 0004 | Sections are weighted horizontal lanes stored in the database | Proposed |
| 0005 | Configuration lives in the database; environment holds secrets and seeds defaults | Proposed |
| 0006 | Render server-side with Pillow to a 1-bit PNG | Proposed |
| 0007 | SQLite and a single application process | Proposed |

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
