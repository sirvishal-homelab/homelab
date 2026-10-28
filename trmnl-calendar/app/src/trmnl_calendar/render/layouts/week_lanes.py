"""Week - lanes: categories down, days across, trips as bars.

Geometry follows docs/04 section 9 and the approved mock-up.
"""

from __future__ import annotations

from datetime import date, timedelta

from trmnl_calendar.domain.models import Occurrence
from trmnl_calendar.domain.recurrence import week_floor
from trmnl_calendar.render import text as T
from trmnl_calendar.render.layouts.base import (
    ALL_DAYS,
    DAY_ABBR,
    PlanContext,
    fmt_range,
    fmt_time,
    title_range,
)
from trmnl_calendar.render.raster import SHAPES
from trmnl_calendar.render.scene import FitReport, Glyph, Line, Rect, Scene, Text

LAYOUT_ID = "week_lanes"
LABEL = "Week - lanes"

MARGIN = 12
GUTTER = 58
HEADER_TOP = 44
BAR_H = 16
BAR_STEP = 19
MAX_BAR_LEVELS = 3
MIN_LANE_PX = 28
GREY = 170


def window(today: date, week_start: str = "sun") -> tuple[date, date]:
    first = week_floor(today, week_start)
    return first, first + timedelta(days=6)


def _visible_days(first: date, mask: int) -> list[date]:
    days = [first + timedelta(days=i) for i in range(7)]
    shown = [d for d in days if mask >> d.weekday() & 1]
    return shown or days


def plan(ctx: PlanContext) -> tuple[Scene, FitReport]:
    W, H = ctx.width, ctx.height
    first, last = window(ctx.today, ctx.week_start)
    days = _visible_days(first, ctx.weekdays if ctx.weekdays else ALL_DAYS)
    col_of = {d: i for i, d in enumerate(days)}
    ncols = len(days)
    cw = (W - 2 * MARGIN - GUTTER) // ncols
    scene, fit = Scene(W, H), FitReport()
    add = scene.items.append
    if (W, H) != (800, 480):
        fit.add("profile_unproven", f"{W} x {H} is not a proven size; the layout is untested here.")

    # today highlight: today, else the next visible day inside the week
    today_col = next((col_of[d] for d in days if d >= ctx.today), None)

    add(Text(MARGIN + 2, 29, title_range(first, last), 22, True))
    y0 = HEADER_TOP
    add(Line(MARGIN, y0 + 22, W - MARGIN, y0 + 22, width=2))

    lanes = [s for s in ctx.sections if s.placement == "lane"]
    banners = [s for s in ctx.sections if s.placement == "banner"]
    banner_y = y0 + 24
    lane_top = banner_y + 17 if banners else y0 + 27
    lane_bot = H - 14

    occ_by = {}
    hidden = 0
    for o in ctx.occurrences:
        if o.date < first or o.date > last:
            continue
        if o.date not in col_of:
            hidden += 1
            continue
        occ_by.setdefault((o.section_id, o.date), []).append(o)

    # header: day names, today fill, banner strip, vertical rules
    for i, d in enumerate(days):
        x = MARGIN + GUTTER + i * cw
        if i == today_col:
            add(Rect(x + 2, y0, cw - 4, 22, fill=GREY))
        add(Text(x + cw // 2, y0 + 16, f"{DAY_ABBR[d.weekday()]} {d.day}", 13, True, anchor="m"))
        if banners:
            names = [o.title for s in banners for o in occ_by.get((s.id, d), [])]
            if names:
                add(Rect(x + 2, banner_y, cw - 4, 15, fill=GREY))
                label, _ = T.ellipsize(", ".join(names), 11, True, cw - 12)
                add(Text(x + cw // 2, banner_y + 12, label, 11, True, anchor="m"))
        if i:
            add(Line(x, banner_y if banners else lane_top, x, lane_bot, dotted=True))

    # lanes
    truncated = overflow = 0
    if not lanes:
        add(Text(W // 2, (lane_top + lane_bot) // 2, "No categories to show", 12, anchor="m"))
    total_weight = sum(s.weight for s in lanes) or 1
    avail = lane_bot - lane_top
    ty = lane_top
    for sec in lanes:
        lh = avail * sec.weight // total_weight
        if lh < MIN_LANE_PX:
            fit.add("lane_too_small", f"Lane '{sec.name}' is only {lh} px tall.", sec.name)
        add(Line(MARGIN, ty, W - MARGIN, ty, dotted=True))
        shape = SHAPES[ctx.sections.index(sec) % len(SHAPES)]
        add(Glyph(MARGIN, ty + 6, shape))
        label, _ = T.ellipsize(sec.name, 11, True, GUTTER - 14)
        add(Text(MARGIN + 11, ty + 14, label, 11, True))

        # trips: one bar per series occurrence, on stacked levels
        groups: dict[tuple[int, date], list[Occurrence]] = {}
        for d in days:
            for o in occ_by.get((sec.id, d), []):
                if o.is_multi_day:
                    groups.setdefault((o.event_id, o.series_start), []).append(o)
        bars = []
        for (_, start), occs in groups.items():
            o0 = occs[0]
            end = start + timedelta(days=o0.day_count - 1)
            a, b = col_of[min(o.date for o in occs)], col_of[max(o.date for o in occs)]
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
        col_levels = [0] * ncols
        for a, b, _s, _e, _o, lv in placed:
            for c in range(a, b + 1):
                col_levels[c] = max(col_levels[c], lv + 1)
        for a, b, start, end, o0, lv in placed:
            x = MARGIN + GUTTER + a * cw + 3
            bw = (b - a + 1) * cw - 6
            y = ty + 3 + lv * BAR_STEP
            add(Rect(x, y, bw, BAR_H, fill=GREY, outline=0))
            cut_l, cut_r = start < first, end > last
            title = ("◂ " if cut_l else "") + o0.title
            if not cut_l and not cut_r:
                with_range = f"{o0.title} - {fmt_range(start, end)}"
                if T.measure(with_range, 12, True) <= bw - 10:
                    title = with_range
            if cut_r:
                title += " ▸"
            title, cut = T.ellipsize(title, 12, True, bw - 10)
            truncated += cut
            add(Text(x + 5, y + 13, title, 12, True))

        # single-day items in each day cell
        for i, d in enumerate(days):
            x = MARGIN + GUTTER + i * cw
            items = [o for o in occ_by.get((sec.id, d), []) if not o.is_multi_day]
            yy = ty + 3 + col_levels[i] * BAR_STEP
            shown = 0
            for o in items:
                label = o.title
                if o.start_time and ctx.options.get("show_times", True):
                    label = f"{fmt_time(o.start_time, ctx.time_format)} {o.title}"
                lines, cut = T.wrap(label, 12, False, cw - 8, 2)
                need = 14 * len(lines)
                last_item = o is items[-1]
                if (yy + need > ty + lh - 12 and not last_item) or yy + need > ty + lh - 2:
                    left = len(items) - shown
                    add(Text(x + cw - 5, ty + lh - 4, f"+{left}", 11, True, anchor="r"))
                    overflow += left
                    break
                truncated += cut
                for k, line in enumerate(lines):
                    add(Text(x + 4, yy + 11 + k * 14, line, 12))
                yy += need + 3
                shown += 1
        ty += lh
    add(Line(MARGIN, lane_bot, W - MARGIN, lane_bot, dotted=True))

    if hidden:
        add(Text(W - MARGIN, H - 3, f"+{hidden} on hidden days", 11, anchor="r"))
        fit.add("hidden_day_events", f"{hidden} events fall on days this view hides.")
    if overflow:
        fit.add("events_overflow", f"{overflow} events do not fit and show as +N.")
    if truncated:
        fit.add("titles_truncated", f"{truncated} titles were shortened.")
    return scene, fit
