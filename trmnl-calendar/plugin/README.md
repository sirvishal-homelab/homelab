# plugin/ - Terminus extension variant (NOT PLANNED)

Decision (2026-10-25, ADR-0003 accepted): the app pushes a rendered image to Terminus's Screens API. A native
Terminus extension (option B in `docs/05-terminus-integration.md`) is not needed and **this directory stays
empty**; no code will be added here.

The only related experiment is **E-1** in ADR-0003: a Terminus *poll extension* that fetches the app's
`/calendar.png` as an image source. It needs no files in this repository (it is configured in the Terminus
UI) and is trialled after the push slice. If E-1 changes the decision, a new ADR says what, if anything,
belongs here.
