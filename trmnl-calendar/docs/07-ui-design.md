# 07 - Editor UI design

> **Partly superseded (2026-10-27).** The target layout below predates the phone-first manage area, people,
> undo, saved views and the live native-size preview (`15`, `17`). It is rewritten at the start of the slice
> that builds the UI (`10-roadmap.md`).

Single page, no framework, server-rendered shell plus a small script. Built in v1; the
screenshots show **v1 only** (no section or layout controls yet).

![v1 desktop](assets/ui-v1-desktop.png)

## Information architecture (target)

```
+------------------------------------------------------+------------------------+
| Calendar for TRMNL                                   |                        |
+------------------------------------------------------+  Add / edit an event   |
| What the screen shows                                |  - Title               |
|  [ device-framed live preview of /calendar.png ]     |  - Section  (select)   |
|  Image address [ http://host:8080/calendar.png ][Copy]  - Date / Ends on      |
|  Terminus: connected, last push 2:05 PM [Push now]   |  - All day / Start/End |
+------------------------------------------------------+  - Location            |
| Screen layout                                        |  - Repeats / Until     |
|  View [Month v]  Week starts [Sun v]  Time [12h v]   |  [Save event]          |
|  [x] Weekends  [x] Section labels  [ ] Hide empty    |                        |
+------------------------------------------------------+  (sticky on desktop)   |
| Sections                                             |                        |
|  Events  height[1] [x]show  up down  Delete          |                        |
|  Work    height[1] [x]show  up down  Delete          |                        |
|  [ new section name ] [Add section]                  |                        |
+------------------------------------------------------+                        |
| Your events (grouped: Repeating, then by month)      |                        |
+------------------------------------------------------+------------------------+
```

Behaviours:
- Layout settings and section edits **autosave on change**, announce the result in an
  `aria-live` region, then refresh the preview and trigger a push.
- The preview always shows what the device will get with no `?view=` (the saved layout).
- The image address field shows the URL **as the browser sees it**; the value Terminus needs
  may differ (see `02-architecture.md`, deployment).

## Principles used in v1

- The one memorable element is the e-paper preview inside a device-style bezel.
- Cool grey page, near-black ink, one blue accent used only for primary actions and focus.
- Ruled rows rather than cards. One corner radius family.
- System-font stack with a legible fallback (no web fonts, so it works offline on a LAN).
- Visible focus rings, labelled controls, errors shown beside fields, table-like data uses
  tabular numbers.
- Event titles are inserted with `textContent` so a title can never inject HTML.

## Open UI questions

- Do sections need colours or icons, or does the 1-bit output make position enough? (OQ-6)
- Is autosave right for the sections manager, or should edits be explicit?
- Is a separate "layout preview" per view needed before saving?
- Mobile: the v1 layout stacks the form above the preview; confirm that is the right order.
