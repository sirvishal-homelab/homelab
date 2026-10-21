# 08 - Open questions and risks

Ask the user the **questions first** (blocking ones marked B). Resolve **risks** by experiment.

## Questions for the user

| ID | Question | Why it matters | |
|----|----------|----------------|-|
| OQ-1 | Which TRMNL model is it, and what firmware? 1-bit or 2-bit display? | Image format, resolution, `model_id`, Terminus version. | B |
| OQ-2 | Is Terminus already installed? Which version, where (Docker? Pi?), what URL and port? | Reachability, API base URL. | B |
| OQ-3 | One device or several? One calendar or several (family, work)? | Profiles, multi-calendar data model. | B |
| OQ-4 | Is the app single-user on a trusted LAN forever, or could it be reached by others? | Authentication scope. | |
| OQ-5 | Repeating events: need to skip or edit single occurrences? Month-end behaviour (skip vs clamp)? | Data model for exceptions. | |
| OQ-6 | Should sections carry icons, patterns or colours? Max number you actually want? | Renderer policy for 5-8 lanes. | |
| OQ-7 | Is a background thread inside the web process acceptable, or do you prefer a separate worker/cron? | Process model, Docker design. | |
| OQ-8 | Do you want import from Google Calendar / ICS feeds, or manual entry only? | Biggest scope driver. | |
| OQ-9 | Preferred database approach and migration tool (stay on SQLite?). | Schema evolution. | |
| OQ-10 | Do you want a public API for other tools, or UI-only? | API versioning and stability. | |
| OQ-11 | Language/stack preference for the real build (Python is the prototype only)? | Everything downstream. | B |
| OQ-12 | Is a native Terminus extension (option B) attractive to you, or is image push preferred? | Whether `plugin/` is built at all. | B |

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
