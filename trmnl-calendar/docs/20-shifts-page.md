# 20. Shifts pages (design note)

Status: **agreed 2026-11-01** (owner), built in the same pull request. Decision record: ADR-0014.

## 1. What this is for

A household with shift work (a rotating Week A / Week B pattern) wants to see, in one place, who is
working on which days, and to mark a day **In**, **Out** or **Day off** in one tap. Her partner wants a
view of **just the work schedule, a month at a glance**. Both are the new **Shifts** tab.

## 2. What exists already (not changed)

- A rotation is a repeat rule on one event (`03`, `17` section 6): 2 to 4 weeks, each a set of weekdays,
  Monday-to-Sunday weeks, Week A being the week that contains the start date.
- Per-day notes **In** / **Out** are `event_day_notes`; a **day off** is a skip in `event_exceptions`.
- `POST /events/{id}/day` sets or clears a note, `POST /events/{id}/skip` and `/unskip` handle days off.

## 3. What was built

**No new tables, no migration, no JSON API change.** Two pages:

- `GET /shifts` (**Weeks**): one card per rotating schedule, titled with its people and event title and
  saying which week of the pattern this week is ("Week B this week"). Each card has a strip as many weeks
  long as the rotation (2 to 4, named A to D) starting with this Monday. Tapping a day opens it in the
  month view. Links: Month view, Edit pattern.
- `GET /shifts/month?e=<event>&y=&m=&day=` (**Month**): the work month of **one schedule**, nothing else
  from the calendar. Chips switch between schedules when there are several; the default is the schedule
  of the person this device chose in "Who are you", else the first. Previous and next month, and Today.
  Below the grid, the chosen day shows its state in words and the buttons: In, Out, Day off for a working
  day, **Restore this day** for a day off, an explanation for a day that is not on the pattern.
  Edit pattern links to the event form.

How a day is drawn (words and shapes, never colour alone, because the e-paper has no colour and colour is
unreliable in any case): a **filled circle** is a working day, a **dashed circle** a day off, a plain circle
a day that is not on the pattern; **in**, **out** and **off** are written under the date; today has a ring.
Each day has a spoken label ("Wed Nov 4, working, out").

A day off wins over an In/Out note on the same day. `unskip` now returns to the page that sent it when it
is given `back` (the Month page and the Shifts pages do); without `back` it still goes to the event.

Navigation: **Shifts** is always a tab (owner's choice): Calendar, Month, Shifts, Screen, More on a phone;
in the top bar on a wide screen.

## 4. Decisions

| # | Question | Decision |
|---|----------|----------|
| S-1 | Shifts as a phone tab: always, or only when a schedule exists? | **Always** (owner, 2026-11-01). |
| S-2 | Edit the Week A / Week B pattern on the Shifts pages? | Not in this slice; the event form does it. Linked as "Edit pattern". |
| S-3 | Show the whole rotation or two weeks? | The whole rotation (2 to 4 weeks). |
| S-4 | A free-text note on a date, besides In / Out / Day off? | **Not now** (owner, 2026-11-01). It would need one small table (`event_day_memos`) and a text box on the day; the first slice does not stand in its way. |

## 5. Tests

`tests/integration/test_shifts_view.py` (states, strip length, day off beats a note, empty week, no
schedules, month grid and panel, schedule chosen by "Who are you", week start does not move the pattern)
and page tests in `test_pages.py` (tab always present, empty pages, In / Out / Day off / Restore return to
the month view, `unskip` without `back`).

## 6. Not in scope

A free-text note per date, pattern editing on these pages, swapping shifts between people, hours or times
per shift, reminders.
