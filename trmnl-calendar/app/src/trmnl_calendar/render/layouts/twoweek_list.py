"""Two weeks - list: two week rows, each day a short list with category symbols, trips as bars.

Geometry follows docs/04 section 9 and the approved mock-up (assets/options/2week-B-flow-trips.png).
"""

from __future__ import annotations

from datetime import date, timedelta

from trmnl_calendar.domain.models import Occurrence
from trmnl_calendar.domain.recurrence import week_floor
from trmnl_calendar.render import text as T
from trmnl_calendar.render.layouts.base import (
    ALL_DAYS,
    DAY_ABBR,
    MONTH_ABBR,
    PlanContext,
    fmt_range,
    fmt_time,
    title_range,
)
from trmnl_calendar.render.raster import SHAPES
from trmnl_calendar.render.scene import FitReport, Glyph, Line, Rect, Scene, Text

LAYOUT_ID = "twoweek_list"
LABEL = "Two weeks - list"
OPTIONS = ("show_times", "show_locations", "show_banners", "show_hidden_marker")

MARGIN = 12
GREY = 170
BAR_H = 15
BAR_STEP = 17
MAX_BAR_LEVELS = 3
LINE_H = 13
HEADER_Y = 52
BODY_TOP = 62
FOOTER_H = 24


def window(today: date, week_start: str = "sun") -> tuple[date, date]:
    first = week_floor(today, week_start)
    return first, first + timedelta(days=13)


def _visible(first: date, mask: int) -> list[date]:
    days = [first + timedelta(days=i) for i in range(7)]
    return [d for d in days if mask >> d.weekday() & 1] or days


def plan(ctx: PlanContext) -> tuple[Scene, FitReport]:
    W, H = ctx.width, ctx.height
    first, last = window(ctx.today, ctx.week_start)
    mask = ctx.weekdays if ctx.weekdays else ALL_DAYS
    scene, fit = Scene(W, H), FitReport()
    add = scene.items.append
    if (W, H) != (800, 480):
        fit.add("profile_unproven", f"{W} x {H} is not a proven size; the layout is untested here.")

    rows = [_visible(first, mask), _visible(first + timedelta(days=7), mask)]
    ncols = len(rows[0])
    cw = (W - 2 * MARGIN) // ncols
    all_visible = {d for r in rows for d in r}
    today_day = next((d for r in rows for d in r if d >= ctx.today), None)

    lanes = [s for s in ctx.sections if s.placement == "lane"]
    banners = (
        {s.id for s in ctx.sections if s.placement == "banner"}
        if ctx.opt("show_banners")
        else set()
    )
    lane_index = {s.id: i for i, s in enumerate(lanes)}
    shape_of = {s.id: SHAPES[ctx.sections.index(s) % len(SHAPES)] for s in lanes}

    by_day: dict[date, list[Occurrence]] = {}
    hidden = 0
    for o in ctx.occurrences:
        if o.date < first or o.date > last:
            continue
        if o.date not in all_visible:
            hidden += 1
            continue
        by_day.setdefault(o.date, []).append(o)

    add(Text(MARGIN + 2, 29, title_range(first, last), 22, True))
    body_bot = H - FOOTER_H
    row_h = (body_bot - BODY_TOP) // 2

    # day-name header
    for i, d in enumerate(rows[0]):
        add(Text(MARGIN + i * cw + cw // 2, HEADER_Y, DAY_ABBR[d.weekday()], 12, True, anchor="m"))
    add(Line(MARGIN, BODY_TOP, W - MARGIN, BODY_TOP, width=2))

    truncated = overflow = 0
    for r, days in enumerate(rows):
        top = BODY_TOP + 2 + r * row_h
        bottom = top + row_h - 2
        col_of = {d: i for i, d in enumerate(days)}
        row_first, row_last = days[0], days[-1]
        if r:
            add(Line(MARGIN, top - 2, W - MARGIN, top - 2, dotted=True))

        # date labels and banner strips
        for i, d in enumerate(days):
            x = MARGIN + i * cw
            label = f"{MONTH_ABBR[d.month - 1]} {d.day}" if d.day == 1 or d == first else str(d.day)
            lw = int(T.measure(label, 13, True))
            if d == today_day:
                add(Rect(x + 2, top, lw + 6, 16, fill=GREY))
            add(Text(x + 5, top + 13, label, 13, True))
            names = [o.title for o in by_day.get(d, []) if o.section_id in banners]
            if names:
                bx = x + 5 + lw + 6
                bw = x + cw - 3 - bx
                if bw > 24:
                    add(Rect(bx, top + 1, bw, 14, fill=GREY))
                    text, cut = T.ellipsize(", ".join(names), 11, True, bw - 6)
                    truncated += cut
                    add(Text(bx + 3, top + 12, text, 11, True))
            if i:
                add(Line(x, top, x, bottom, dotted=True))

        # trips: bars in a strip under the date labels
        groups: dict[tuple[int, date], list[Occurrence]] = {}
        for d in days:
            for o in by_day.get(d, []):
                if o.is_multi_day and o.section_id not in banners:
                    groups.setdefault((o.event_id, o.series_start), []).append(o)
        bars = []
        for (_, start), occs in groups.items():
            o0 = occs[0]
            end = start + timedelta(days=o0.day_count - 1)
            a = col_of[min(o.date for o in occs)]
            b = col_of[max(o.date for o in occs)]
            bars.append((a, b, start, end, o0))
        bars.sort(key=lambda t: (t[0], -(t[1] - t[0]), t[4].title.lower(), t[4].event_id))
        level_end: list[int] = []
        placed = []
        for a, b, start, end, o0 in bars:
            for lv, e in enumerate(level_end):
                if a > e:
                    level_end[lv] = b
                    break
            else:
                lv = len(level_end)
                level_end.append(b)
            if lv >= MAX_BAR_LEVELS:
                overflow += 1
                continue
            placed.append((a, b, start, end, o0, lv))
        strip_y = top + 19
        for a, b, start, end, o0, lv in placed:
            x = MARGIN + a * cw + 2
            bw = (b - a + 1) * cw - 4
            y = strip_y + lv * BAR_STEP
            add(Rect(x, y, bw, BAR_H, fill=GREY, outline=0))
            cut_l, cut_r = start < row_first, end > row_last
            title = ("◂ " if cut_l else "") + o0.title
            if not cut_l and not cut_r:
                with_range = f"{o0.title} · {fmt_range(start, end)}"
                if T.measure(with_range, 11, True) <= bw - 10:
                    title = with_range
            if cut_r:
                title += " ▸"
            title, cut = T.ellipsize(title, 11, True, bw - 10)
            truncated += cut
            add(Text(x + 4, y + 12, title, 11, True))
        list_top = strip_y + (len(level_end[:MAX_BAR_LEVELS]) * BAR_STEP) + 2

        # day lists
        for i, d in enumerate(days):
            x = MARGIN + i * cw
            items = sorted(
                (o for o in by_day.get(d, []) if not o.is_multi_day and o.section_id in lane_index),
                key=lambda o: (
                    lane_index[o.section_id],
                    not o.all_day,
                    o.start_time or "",
                    o.title.lower(),
                    o.event_id,
                ),
            )
            yy = list_top
            shown = 0
            for o in items:
                label = o.title
                if ctx.opt("show_locations") and o.location:
                    label = f"{label} @ {o.location}"
                if o.start_time and ctx.opt("show_times"):
                    label = f"{fmt_time(o.start_time, ctx.time_format)} {label}"
                lines, cut = T.wrap(label, 11, False, cw - 18, 3)
                need = LINE_H * len(lines)
                last_item = o is items[-1]
                if (yy + need > bottom - LINE_H and not last_item) or yy + need > bottom:
                    left = len(items) - shown
                    add(Text(x + cw - 4, bottom - 1, f"+{left}", 11, True, anchor="r"))
                    overflow += left
                    break
                truncated += cut
                add(Glyph(x + 5, yy + 3, shape_of[o.section_id], 7))
                for k, line in enumerate(lines):
                    add(Text(x + 15, yy + 10 + k * LINE_H, line, 11))
                yy += need + 2
                shown += 1

    # footer: rule, category legend, hidden-day note
    add(Line(MARGIN, body_bot, W - MARGIN, body_bot, dotted=True))
    lx = MARGIN + 2
    for s in lanes:
        add(Glyph(lx, H - 15, shape_of[s.id], 8))
        name, _ = T.ellipsize(s.name, 11, True, 110)
        add(Text(lx + 12, H - 7, name, 11, True))
        lx += 12 + int(T.measure(name, 11, True)) + 14
        if lx > W - 200:
            break
    if hidden and ctx.opt("show_hidden_marker"):
        add(Text(W - MARGIN, H - 7, f"+{hidden} on hidden days", 11, anchor="r"))
    if hidden:
        fit.add("hidden_day_events", f"{hidden} events fall on days this view hides.")
    if overflow:
        fit.add("events_overflow", f"{overflow} events do not fit and show as +N.")
    if truncated:
        fit.add("titles_truncated", f"{truncated} titles were shortened.")
    return scene, fit
