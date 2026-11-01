# 20. Shifts page (design note, for the owner's review)

Status: **proposed 2026-11-01**. No code yet. Decision record: ADR-0014.

## 1. What this is for

A household with shift work (a rotating Week A / Week B pattern) wants to see, in one place, who is
working on which days for the next few weeks, and to mark a day **In**, **Out** or **Day off** in one tap.
Today that is possible, but only by finding the "Work" event or scrolling the home list. The mockups gave
this its own **Shifts** page.

## 2. What exists already (no change needed)

- A rotation is a repeat rule on one event (`03`, `17` section 6): 2 to 4 weeks, each a set of weekdays,
  Monday-to-Sunday weeks, Week A being the week that contains the start date.
- Per-day notes **In** / **Out** are `event_day_notes`; a **day off** is a skip in `event_exceptions`.
- `POST /events/{id}/day` sets or clears a note, `POST /events/{id}/skip` and `/unskip` handle days off.
- The home list already shows the three buttons under a scheduled day.

## 3. Proposal

`GET /shifts`: a read-and-quick-edit overview. **No new tables, no migration, no API change.**

```
Shifts                         [ Everyone | Alex | Sam ]

Sam: Work        Week B this week          [Edit pattern]
  Week of Mon 26 Oct
   M   T   W   T   F   S   S
  26  27  28  29  30  31   1
           out  in
  Week of Mon 2 Nov
   2   3   4   5   6   7   8
  [w]  off
  Tap a working day for In / Out / Day off.
```

- One card per rotating schedule (a non-deleted event with a rotation), titled with its people and title.
  Filter chips by person, like the home list.
- A strip of **as many weeks as the rotation has** (2 to 4), starting with the Monday of this week, so the
  whole pattern is visible. Each day shows its date and one of: not a working day (blank), working and not
  decided, **in**, **out**, **day off**. The state is shown with text, not colour alone, because colour is
  not available on the e-paper and is hard to rely on in any case.
- Tapping a working day opens the same In / Out / Day off buttons the home list uses (reused, posting
  back to `/shifts`). A day off shows **Restore** (the existing unskip).
- The card header says which week of the pattern this week is ("Week B this week"), named A, B, C, D.
- **Edit pattern** and **Add a schedule** link to the existing event form. The pattern editor stays there
  in this slice (see open question 2).

## 4. How it is built

- `domain/recurrence.py` gets one pure function, `rotation_week(event, day) -> int | None` (which week of
  the pattern a date falls in), refactored out of the code `occurrences()` already uses. Unit-tested,
  including a three-week rotation and a week that is empty.
- `views.shifts_context(conn, now, person)` builds the cards: events with a rotation, their skips and
  notes for the window, and a cell state per day. It needs the pattern days **and** the skips, because a
  day off is a scheduled day that was skipped.
- `web/pages_shifts.py` with `GET /shifts`; the existing day and skip routes accept `back=/shifts`.
  `template shifts.html`; a little CSS beside the existing `.dayactions` rules.
- Navigation: **Shifts** appears in the phone bar as a fifth tab only when at least one rotating schedule
  exists (one cheap query in `chrome()`); otherwise it is a row on the More page. On a wide screen it is in
  the top bar. See open question 1.
- Tests: unit (week index, cell states, empty week, skipped day), integration (page with and without a
  schedule, tapping In / Out / Day off / Restore returns to `/shifts`, person filter, the tab appears only
  when a schedule exists).
- No effect on the e-paper picture beyond what a status change already does (it already changes it).

## 5. Open questions for the owner

| # | Question | Recommendation |
|---|----------|----------------|
| S-1 | A fifth phone tab for Shifts, always, or only when a schedule exists? | Only when one exists; always on More. Keeps the bar to four tabs for households without shifts. |
| S-2 | Edit the Week A / Week B pattern on the Shifts page itself? | Not in the first slice. It needs a new save route and conflict handling; the event form already does it. Add later if the extra tap bothers you. |
| S-3 | Show the whole rotation length, or always two weeks? | Whole rotation (2 to 4 weeks), so a three-week pattern is not hidden. |
| S-4 | Should In / Out / Day off also be offered from the Month page's day list? | They already are for scheduled days; leave as is. |

## 6. Not in scope

Pattern editing on this page, swapping shifts between people, hours or times per shift, reminders,
and anything that changes how rotations are stored.
