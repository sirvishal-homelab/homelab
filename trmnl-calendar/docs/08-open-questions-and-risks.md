# 08 - Open questions and risks

Ask the user the **questions first** (blocking ones marked B). Resolve **risks** by experiment.

## Questions for the user

| ID | Question | Why it matters | |
|----|----------|----------------|-|
| OQ-1 | Which TRMNL model is it, and what firmware? 1-bit or 2-bit display? | Image format, resolution, `model_id`, Terminus version. **Answered: TRMNL OG, firmware 1.8.6, model `og_plus`, 2-bit, 800x480 (`14`).** |
| OQ-2 | Is Terminus already installed? Which version, where (Docker? Pi?), what URL and port? | Reachability, API base URL. **Answered: Docker on an Ubuntu container, `http://<server>:2300`, rolling build ahead of 0.78.0 (`13`).** |
| OQ-3 | One device or several? One calendar or several (family, work)? | Profiles, multi-calendar data model. **Answered: one device and one calendar; schema ready for more (A-1).** |
| OQ-4 | Is the app single-user on a trusted LAN forever, or could it be reached by others? | Authentication scope. | **Answered: LAN only, no login in v1; optional shared password later (`02` section 9).** |
| OQ-5 | Repeating events: need to skip or edit single occurrences? Month-end behaviour (skip vs clamp)? | Data model for exceptions. | **Answered: skips, clamp at month end, nth weekday under More options (`03`, ADR-0010).** |
| OQ-6 | Should sections carry icons, patterns or colours? Max number you actually want? | Renderer policy for 5-8 lanes. | **Answered for now: categories (max 8), symbols and banner placement; no colours (`16`).** |
| OQ-7 | Is a background thread inside the web process acceptable, or do you prefer a separate worker/cron? | Process model, Docker design. | **Answered: one process, one worker, state in the database (ADR-0007).** |
| OQ-8 | Do you want import from Google Calendar / ICS feeds, or manual entry only? | Biggest scope driver. | **Answered: manual entry in v1; ICS later (A-3).** |
| OQ-9 | Preferred database approach and migration tool (stay on SQLite?). | Schema evolution. | **Answered: SQLite with numbered migrations and a pre-migration backup (`03`).** |
| OQ-10 | Do you want a public API for other tools, or UI-only? | API versioning and stability. | **Answered: UI-only API, versioned, PATCH not PUT (`06`).** |
| OQ-11 | Language/stack preference for the real build (Python is the prototype only)? | Everything downstream. **Answered: Python (A-2); FastAPI stack proposed in ADR-0009.** |
| OQ-12 | Is a native Terminus extension (option B) attractive to you, or is image push preferred? | Whether `plugin/` is built at all. **Answered: image push (ADR-0003); E-1 trial before it is final.** |

Status column added 2026-10-27: every question is now answered; the answer is appended in bold.

## Risks

| ID | Risk | Likelihood | Impact | Mitigation |
|----|------|-----------|--------|------------|
| R-1 | Terminus API differs from assumptions (field names, auth). | Medium | High | Spike against a real Terminus first (see roadmap). Keep the client behind an interface. |
| R-2 | Terminus is beta; API changes break pushes. | Medium | Medium | Pin a Terminus version; contract test the client. |
| R-3 | Terminus refuses LAN image URLs or re-processes our 1-bit image badly. | Low-medium | High | Spike T-5, T-6. Fallback: preprocessed image or option B. |
| R-4 | Docker font mismatch changes layouts. | High (likely) | Medium | Bundle a font file; snapshot tests inside the container. |
| R-5 | In-process worker plus multi-worker server causes duplicate pushes. | Medium | Low | Single process, or separate publisher. |
| R-6 | No auth on the app and credentials in the environment. | Low on a LAN | Medium | Document; consider a shared token for write APIs. |
| R-7 | Month view unreadable with many sections. | Medium | Medium | Cap lanes, abbreviate, or collapse empty lanes. |
| R-8 | Naive local dates and a server `TZ` setting give a wrong "today" if misconfigured. | Medium | Low-medium | Make timezone an explicit, validated setting; show it in the UI. |
| R-9 | Hand-rolled migrations drift. | Medium | Medium | Introduce versioned migrations before the next schema change. |
| R-10 | The prototype is mistaken for a finished base. | Medium | Medium | `STATUS.md`; treat docs as authoritative. |

## Known defects in the prototype

1. `app.py`, templates and JS do not use sections, settings or the publisher.
2. Dockerfile uses 2 workers and the wrong font package; never built.
3. `PUT /api/events/<id>` resets omitted fields (bit me during testing).
4. `publisher.py` has no tests; `_authed` fallback logic is unproven.
5. The publisher's `request_push` reason is overwritten by the latest call (fine, but untidy).
6. No automated tests anywhere; all checks were ad hoc scripts and screenshots.


## Risk status update 2026-10-27

| ID | Status |
|----|--------|
| R-1 | Retired for the verified calls (`14`); remains for upgrades: keep the contract tests and run the spike after each Terminus change |
| R-2 | Open and real: Terminus pushed an unattended update on 2026-10-23; pin it (`13`, P-3) |
| R-3 | Retired: LAN URLs work; `preprocessed: true` stores our PNG unchanged |
| R-4 | Mitigated by design: bundled fonts and goldens in the container (`04`, `18`) |
| R-5 | Retired by ADR-0007: one process, one worker |
| R-6 | Open: no login plus credentials in the environment; see `02` section 9 |
| R-7 | Mitigated by the fit check and policy (`04` section 8) |
| R-8 | Mitigated by an explicit calendar timezone and floating local time (ADR-0010) |
| R-9 | Mitigated by numbered migrations (`03`) |
| R-10 | Mitigated by STATUS.md and the docs |
| R-11 | New: render-time text would defeat push-only-on-change; fixed by ADR-0006 and `04` section 6 |
| R-12 | New: Terminus deletes the previous image on re-render (upstream issue 384); push on change only keeps it rare |
| R-13 | New: the public repo exposes home-network detail in `13` and `spikes/` (`12`) |
| R-14 | New 2026-10-29: the online holiday lookup was written without calling the live service, and its v3 ends 2027-01-31. Open until one real lookup (for example Canada, with Nova Scotia) is checked, then move to v4 (ADR-0011) |
