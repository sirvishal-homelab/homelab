# TRMNL Calendar

A self-hosted calendar for a [TRMNL](https://usetrmnl.com) e-paper display, served through a
local [Terminus](https://github.com/usetrmnl/terminus) (BYOS) server on a home network.

You enter events in a web page. The app stores them, draws them as a black-and-white image
(week, two-week or month layouts, with user-defined rows such as *Events / Work / Lunch / Dinner*
inside each day), and pushes that image to Terminus so the device can show it.

> **Status: design phase.** A working prototype exists, but it is partly unfinished and
> partly untested. The next step is to settle the architecture and design documents before
> writing production code. See [`AGENT_PROMPT.md`](AGENT_PROMPT.md).

## Repository layout

```
.
├── AGENT_PROMPT.md            Hand-off prompt for the next agent (start here)
├── README.md                  This file
├── docs/                      Design documents (the source of truth for design)
│   ├── 00-overview.md         Goals, scope, glossary
│   ├── 01-requirements.md     Functional and non-functional requirements
│   ├── 02-architecture.md     Context, components, deployment, sequences (Mermaid)
│   ├── 03-data-model.md       ER diagram, tables, settings, migrations
│   ├── 04-rendering-design.md Layout engine, section lanes, wireframes, sample images
│   ├── 05-terminus-integration.md  What is known/unknown about Terminus, options compared
│   ├── 06-app-api.md          HTTP API (existing and planned)
│   ├── 07-ui-design.md        Editor UI information architecture and wireframe
│   ├── 08-open-questions-and-risks.md
│   ├── 09-references.md       Every source used, and what still needs reading
│   ├── 10-roadmap.md          Suggested design-first sequence
│   ├── 11-prototype-code-tour.md  Map of the prototype code with key excerpts
│   ├── decisions/             Architecture Decision Records (ADR-0001 ...)
│   └── assets/                Rendered sample screens and v1 UI screenshots
├── app/                       FUTURE: the calendar application (proposed structure only)
├── plugin/                    FUTURE: Terminus extension variant (proposed structure only)
└── reference/
    └── prototype/             Code written during the first session. Reference, not a base.
```

## How to read this repo

1. `AGENT_PROMPT.md` for the situation in one page.
2. `docs/00` to `docs/05` for the design as it stands.
3. `docs/08` and `docs/09` for what is unresolved and where to verify it.
4. `reference/prototype/` only if you want to see what was built. See its `STATUS.md`.

Diagrams are written in [Mermaid](https://mermaid.js.org/), which GitHub renders in Markdown,
and in plain ASCII where layout matters.
