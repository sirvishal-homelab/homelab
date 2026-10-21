# Hand-off prompt: TRMNL Calendar (design phase)

Paste everything below the line into the new project as the first message, and attach this
repository (or its `docs/` folder and `reference/prototype/`).

---

You are taking over a project from a previous working session. **Your job right now is
design and architecture, not coding.** Do not write production code until the design has been
reviewed and approved by me. Throwaway scripts to *verify facts* (for example calling an API
with curl) are fine and encouraged.

## The project in one paragraph

I own a TRMNL e-paper display and run, or am about to run, a local **Terminus** (the TRMNL
"BYOS" server) on my home network. I want a **calendar** on the display. I enter events in a
web page; they are stored in a database and drawn as a black-and-white **image** that Terminus
serves to the device. The image must be customisable: a **week, two-week or month** view, and
each calendar cell divided into user-defined **sections** such as *Events, Work, Lunch, Dinner*.
All of that configuration must be **modular and stored in the database** so nothing is lost.
The app should update Terminus on a schedule and when I change something.

## What exists

A prototype in `reference/prototype/` (Python, Flask, SQLite, Pillow). **Treat it as reference
material, not a foundation.** Its true state, from `reference/prototype/STATUS.md`:

| Part | State |
|------|-------|
| Data layer (`store.py`): events, recurrence, sections, settings, migration | Written; v1 portion tested through the API; sections/settings only exercised indirectly |
| Renderer (`render.py`): week, two-week, month with sections, plus agenda and split | Written; checked **by eye** for several settings; no automated tests |
| Terminus push client (`publisher.py`) | Written; **never run** |
| Web app (`app.py`), UI (`templates/`, `static/`), Docker files | **Stale, v1 only.** No section manager, settings UI, or push wiring. Docker uses 2 workers and probably the wrong font package |

Sample rendered screens are in `docs/assets/`. Read `docs/11-prototype-code-tour.md` for a map
of the code with key excerpts.

## Decisions made so far (all *proposed*, none signed off)

Full text in `docs/decisions/`. In brief:

1. **Design first**; the prototype is reference only (ADR-0001).
2. A **separate app owns the data and the drawing**; Terminus is only the delivery channel (ADR-0002).
3. **Push** the image: log in to Terminus, create one screen, then `PATCH` it on changes and on
   a timer, because from the docs I read, Terminus fetches an image URL only when the screen is
   created or updated (ADR-0003). **Unverified.**
4. **Sections are weighted horizontal lanes** stored in the DB (name, order, relative height
   1-4, visible); each event belongs to one section (ADR-0004).
5. **Config in the database**; environment variables hold secrets and infrastructure and only
   *seed* defaults (ADR-0005).
6. **Pillow renders to a 1-bit PNG** with layouts as registered modules (ADR-0006).
7. **SQLite, single application process**, because the publisher is an in-process thread (ADR-0007).

## Things I want you to be sceptical about

- **Terminus facts came from reading docs, never from running Terminus.** `docs/05` lists nine
  unknowns (T-1 to T-9) with how to verify each: request/response bodies, port, auth header
  (bare token vs `Bearer`), correct `model_id`, how an "unprocessed" 1-bit PNG is converted,
  whether LAN image URLs are allowed, and more. Resolve these before finalising the design.
- The sandbox the prototype was built in had **no network access** to Terminus, so the push
  code is entirely speculative.
- **Font risk:** the renderer prefers DejaVu Sans *Condensed*; the Dockerfile installs a package
  that probably lacks it, so Docker output would differ from the samples. Verify; bundling a
  font is the likely answer.
- `PUT /api/events/<id>` is a full replace (omitted fields reset). Reconsider.
- Month view was only tested with up to four sections; the data model allows eight.
- The "reference" prototype contains my first guesses, not requirements. Challenge them.

## How I want you to work

1. **Read first, in this order:** this prompt; `README.md`; `docs/00` to `docs/05`;
   `docs/08` and `docs/09`; then `reference/prototype/STATUS.md`. Look at the code only after.
2. **Verify before you trust.** Open the primary sources listed in `docs/09-references.md`
   (Terminus `README`, `doc/api.adoc`, `doc/extensions.adoc`, and the other `doc/` files) and
   re-read them in full. Do not rely on the summaries in this repo for anything rated below
   **High** confidence in `docs/05`. Record what you confirm or correct.
3. **Ask me the blocking questions first** (`docs/08`, marked B: device model and bit depth,
   Terminus version and location, one or many devices/calendars, stack preference, image push
   vs native Terminus extension). Ask in small groups, not all at once, and give your
   recommendation with each.
4. **Be honest about uncertainty.** Say what you verified, what you inferred, and what you
   could not check. If you find the prototype or these docs are wrong, say so plainly.
5. **Update the docs as the design evolves.** The `docs/` folder is the source of truth. New
   decisions become ADRs. Keep diagrams in Mermaid or ASCII so they live in Git.
6. **Do not start `app/` or `plugin/`.** Only propose their final structure
   (`app/README.md` and `plugin/README.md` hold my first-draft proposals).

## Deliverables I expect from the design phase

1. Answers to the blocking questions, recorded in `docs/08`.
2. A **Terminus spike log**: evidence (commands and responses, or notes) resolving T-1 to T-6,
   with the confidence column in `docs/05` updated.
3. A decision on **option A vs B vs C** in `docs/05`, with the reasoning, as an accepted ADR.
4. Revised `docs/02` (architecture), `03` (data model), `04` (rendering and the layout-module
   interface), `06` (API), and a settled folder structure for `app/` and `plugin/`.
5. A **test strategy** (fixture data, fixed-time image snapshots, recurrence tests, a fake
   Terminus for contract tests) before any implementation.
6. A short **implementation plan** in vertical slices (see `docs/10-roadmap.md`), each demoable.

## Constraints and preferences

- Runs on my **home network**; no internet exposure. No login in v1, but say clearly if you
  think that is a mistake.
- Prefer few moving parts and something that runs on a small always-on machine.
- The display is **black and white**. Anything drawn must survive thresholding; avoid small
  white-on-black text and reliance on grey (see lessons in `docs/04`).
- Keep explanations short and direct. Use diagrams where they help; prose otherwise.
- I may buy a paid plan and continue here, so write things down in the repo rather than
  relying on conversation memory.

## Where things are

```
AGENT_PROMPT.md   this file
README.md         repo map
docs/             design documents, ADRs, rendered samples
app/              future application (proposal only)
plugin/           future Terminus extension (proposal only, optional)
reference/prototype/   first-session code, labelled with STATUS.md
```

## Your first reply should contain

1. A 5-line summary of the project in your own words, to confirm you understood.
2. The three or four most important **risks** you see in the current design.
3. Your **first group of blocking questions** for me.
4. The **order in which you will verify the Terminus unknowns**.

Then wait for my answers before changing any design documents.
