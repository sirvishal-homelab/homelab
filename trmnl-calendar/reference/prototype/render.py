"""Draws the calendar as a 1-bit PNG sized for a TRMNL e-paper display.

Layout choices (view, week start, sections, ...) come from the `cfg` and `sections`
arguments, which the app loads from the database on every request.
"""
import calendar as pycal
import os
from datetime import date, timedelta
from functools import lru_cache
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont

import store

VIEWS = store.VIEWS
AGENDA_DAYS = 60

_FONT_DIRS = ("/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/dejavu",
              "/usr/share/fonts/TTF", "/Library/Fonts", "C:/Windows/Fonts")
_REGULAR = ("DejaVuSansCondensed.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf", "arial.ttf")
_BOLD = ("DejaVuSansCondensed-Bold.ttf", "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf", "arialbd.ttf")


@lru_cache(maxsize=None)
def _font_path(bold):
    override = os.environ.get("FONT_BOLD" if bold else "FONT_REGULAR")
    if override and os.path.exists(override):
        return override
    for folder in _FONT_DIRS:
        for name in (_BOLD if bold else _REGULAR):
            path = os.path.join(folder, name)
            if os.path.exists(path):
                return path
    return None


@lru_cache(maxsize=None)
def load_font(size, bold=False):
    path = _font_path(bold)
    if path:
        return ImageFont.truetype(path, size)
    return ImageFont.load_default(size)


# ------------------------------------------------------------------ helpers

def fmt_time(hhmm, fmt):
    h, m = int(hhmm[:2]), hhmm[3:]
    if fmt == "24":
        return f"{h:02d}:{m}"
    return f"{h % 12 or 12}:{m} {'AM' if h < 12 else 'PM'}"


def short_time(hhmm, fmt):
    """Compact form for narrow cells: 9a, 9:30a, 17:30."""
    h, m = int(hhmm[:2]), hhmm[3:]
    if fmt == "24":
        return f"{h:02d}:{m}"
    suffix = "a" if h < 12 else "p"
    return f"{h % 12 or 12}{suffix}" if m == "00" else f"{h % 12 or 12}:{m}{suffix}"


def first_weekday(cfg):
    return 6 if cfg["week_start"] == "sun" else 0     # calendar module: Monday=0 ... Sunday=6


def week_start_of(d, first):
    return d - timedelta(days=(d.weekday() - first) % 7)


def column_offsets(first, show_weekends):
    return [i for i in range(7) if show_weekends or (first + i) % 7 < 5]


def fmt_range(a, b):
    ma, mb = pycal.month_abbr[a.month], pycal.month_abbr[b.month]
    if a.year != b.year:
        return f"{ma} {a.day}, {a.year} – {mb} {b.day}, {b.year}"
    if a.month == b.month:
        return f"{ma} {a.day} – {b.day}, {b.year}"
    return f"{ma} {a.day} – {mb} {b.day}, {b.year}"


class Canvas:
    def __init__(self, w, h):
        self.w, self.h = w, h
        self.s = h / 480.0                      # layouts are designed at 800x480
        self.im = Image.new("L", (w, h), 255)
        self.d = ImageDraw.Draw(self.im)

    def px(self, v):
        return int(round(v * self.s))

    def font(self, size, bold=False):
        return load_font(max(8, self.px(size)), bold)

    def text(self, xy, s, size, bold=False, fill=0, anchor="la"):
        self.d.text(xy, s, font=self.font(size, bold), fill=fill, anchor=anchor)

    def fit(self, s, font, max_w):
        if max_w <= 0:
            return ""
        if font.getlength(s) <= max_w:
            return s
        while s and font.getlength(s + "…") > max_w:
            s = s[:-1]
        return (s.rstrip() + "…") if s else ""

    def hline(self, x0, x1, y, dotted=True, width=1, gap=3):
        y = int(y)
        if dotted:
            for x in range(int(x0), int(x1), gap):
                self.d.point((x, y), fill=0)
        else:
            self.d.line([(int(x0), y), (int(x1), y)], fill=0, width=max(1, width))

    def vline(self, x, y0, y1):
        x = int(x)
        for y in range(int(y0), int(y1), 3):
            self.d.point((x, y), fill=0)

    def png(self):
        bw = self.im.point(lambda p: 255 if p >= 140 else 0).convert("1")
        buf = BytesIO()
        bw.save(buf, format="PNG", optimize=True)
        return buf.getvalue()


# ----------------------------------------------------------------- sections

def bucket_by_section(entries, ids, default_id):
    out = {}
    for e in entries:
        sid = e["section_id"] if e["section_id"] in ids else default_id
        out.setdefault(sid, []).append(e)
    return out


def pick_sections(sections, by_day, dates, cfg):
    """Sections to draw: only visible ones, and optionally only those with events in range."""
    ids = {s["id"] for s in sections}
    default_id = sections[0]["id"]
    chosen = [s for s in sections if s["visible"]]
    if cfg["hide_empty_sections"]:
        used = set()
        for d in dates:
            for sid in bucket_by_section(by_day.get(d, []), ids, default_id):
                used.add(sid)
        chosen = [s for s in chosen if s["id"] in used] or chosen[:1]
    return chosen or sections[:1]


def lane_heights(chosen, total):
    weights = [s["weight"] for s in chosen]
    unit = total / float(sum(weights))
    heights, y = [], 0.0
    for wgt in weights:
        heights.append((y, wgt * unit))
        y += wgt * unit
    return heights


def wrap_text(c, text, font, max_w, max_lines):
    """Greedy word wrap; the last allowed line is truncated with an ellipsis."""
    words, lines, cur = text.split(), [], ""
    for word in words:
        trial = (cur + " " + word).strip()
        if font.getlength(trial) <= max_w or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = c.fit(lines[-1] + " " + " ".join(words[len(" ".join(lines).split()):]), font, max_w)
    return [c.fit(line, font, max_w) for line in lines] or [""]


def draw_lane_lines(c, x, y, w, h, entries, cfg, size=12, line_h=15, show_time=True, wrap=2):
    """Several lines per lane (week views). Titles wrap; all-day entries get a solid bar."""
    font = c.font(size)
    lh = c.px(line_h)
    bar = max(3, c.px(3))
    indent = c.px(7)
    budget = max(1, int(h // lh))
    blocks = []
    for e in entries:
        label = e["title"]
        if show_time and e["start_time"]:
            label = f"{short_time(e['start_time'], cfg['time_format'])} {label}"
        blocks.append((e, wrap_text(c, label, font, w - indent, wrap)))
    total = sum(len(lines) for _, lines in blocks)
    shown, used = [], 0
    reserve = 0 if total <= budget else 1
    for e, lines in blocks:
        if used + len(lines) > budget - reserve:
            lines = lines[:max(0, budget - reserve - used)]
            if not lines:
                break
        shown.append((e, lines))
        used += len(lines)
    for e, lines in shown:
        if e["all_day"]:
            c.d.rectangle([x, y + 1, x + bar - 1, y + lh * len(lines) - 2], fill=0)
        for line in lines:
            c.d.text((x + indent, y), line, font=font, fill=0)
            y += lh
    hidden = len(blocks) - len(shown)
    if hidden > 0:
        c.text((x + indent, y), f"+{hidden} more", size, bold=True)


def draw_lane_oneline(c, x, y, w, h, entries, size):
    """One line per lane (month cells): first entry, plus a +N count."""
    if not entries:
        return
    font, bold = c.font(size), c.font(size, True)
    bar = max(3, c.px(3))
    indent = c.px(6)
    first = entries[0]
    suffix = f" +{len(entries) - 1}" if len(entries) > 1 else ""
    suffix_w = bold.getlength(suffix) if suffix else 0
    mid = y + h / 2.0
    if first["all_day"]:
        c.d.rectangle([x, mid - h * 0.38, x + bar - 1, mid + h * 0.38], fill=0)
    label = c.fit(first["title"], font, w - indent - suffix_w)
    c.d.text((x + indent, mid), label, font=font, fill=0, anchor="lm")
    if suffix:
        c.d.text((x + indent + font.getlength(label), mid), suffix, font=bold, fill=0, anchor="lm")


# --------------------------------------------------------------------- grid

def draw_grid(c, x, y, w, h, mode, today, by_day, all_sections, cfg):
    """Week, two-week and month layouts with sections as rows inside each day."""
    first = first_weekday(cfg)
    offsets = column_offsets(first, cfg["show_weekends"])
    ncol = len(offsets)

    if mode == "month":
        weeks = pycal.Calendar(first).monthdatescalendar(today.year, today.month)
        title = f"{pycal.month_name[today.month]} {today.year}"
    else:
        ws = week_start_of(today, first)
        count = 1 if mode == "week" else 2
        weeks = [[ws + timedelta(days=7 * k + i) for i in range(7)] for k in range(count)]
        title = fmt_range(weeks[0][offsets[0]], weeks[-1][offsets[-1]])

    dates = [d for wk in weeks for d in wk]
    chosen = pick_sections(all_sections, by_day, dates, cfg)
    ids = {s["id"] for s in all_sections}
    default_id = all_sections[0]["id"]

    title_h = c.px(36)
    c.text((x, y), title, 24, bold=True)
    if mode == "month" and cfg["show_section_labels"] and len(chosen) > 1:
        legend_font = c.font(11)
        title_w = c.font(24, True).getlength(title)
        text = "Top to bottom: " + ", ".join(s["name"] for s in chosen)
        text = c.fit(text, legend_font, w - title_w - c.px(24))
        c.d.text((x + w, y + c.px(12)), text, font=legend_font, fill=0, anchor="ra")

    top = y + title_h
    avail = h - title_h
    pad = c.px(4)

    if mode == "month":
        head_h = c.px(20)
        col_w = w / float(ncol)
        for i, off in enumerate(offsets):
            name = pycal.day_abbr[(first + off) % 7]
            c.text((x + col_w * i + col_w / 2, top + head_h - c.px(5)), name, 13, bold=True, anchor="ms")
        c.hline(x, x + w, top + head_h, dotted=False, width=c.px(2))
        body_top = top + head_h
        cell_h = (avail - head_h) / len(weeks)
        num_w = c.px(22)
        for r, wk in enumerate(weeks):
            cy = body_top + cell_h * r
            if r > 0:
                c.hline(x, x + w, cy)
            lanes = lane_heights(chosen, cell_h - c.px(2))
            weights = [s["weight"] for s in chosen]
            min_lane = (cell_h / c.s - 2) * min(weights) / float(sum(weights))
            lane_size = max(9, min(12, int(min_lane * 0.72)))
            for i, off in enumerate(offsets):
                day = wk[off]
                if day.month != today.month:
                    continue
                cx = x + col_w * i
                label = str(day.day)
                nfont = c.font(13, True)
                if day == today:
                    c.d.rectangle([cx + pad, cy + pad, cx + pad + c.px(19), cy + pad + c.px(18)], fill=0)
                    c.d.text((cx + pad + c.px(9.5), cy + pad + c.px(9)), label, font=nfont, fill=255, anchor="mm")
                else:
                    c.d.text((cx + pad + c.px(9.5), cy + pad + c.px(9)), label, font=nfont, fill=0, anchor="mm")
                buckets = bucket_by_section(by_day.get(day, []), ids, default_id)
                lx, lw = cx + num_w + pad, col_w - num_w - pad * 2
                for k, (s, (ly, lh)) in enumerate(zip(chosen, lanes)):
                    if k > 0:
                        c.hline(lx, lx + lw, cy + ly, gap=6)
                    draw_lane_oneline(c, lx, cy + ly, lw, lh, buckets.get(s["id"], []), lane_size)
        return

    # week / two weeks
    show_labels = cfg["show_section_labels"]
    label_font = c.font(11, True)
    gutter = 0
    if show_labels:
        widest = max(label_font.getlength(s["name"]) for s in chosen)
        gutter = int(min(widest + c.px(12), w * 0.16))
    col_w = (w - gutter) / float(ncol)
    head_h = c.px(26)
    gap = c.px(10)
    block_h = (avail - gap * (len(weeks) - 1)) / len(weeks)

    for b, wk in enumerate(weeks):
        by = top + b * (block_h + gap)
        for i, off in enumerate(offsets):
            day = wk[off]
            cx = x + gutter + col_w * i
            label = f"{pycal.day_abbr[day.weekday()]} {day.day}"
            if day == today:
                c.d.rectangle([cx + 1, by, cx + col_w - 1, by + head_h - c.px(3)], fill=0)
                c.text((cx + col_w / 2, by + head_h / 2 - c.px(2)), label, 14, bold=True, fill=255, anchor="mm")
            else:
                c.text((cx + col_w / 2, by + head_h / 2 - c.px(2)), label, 14, bold=True, anchor="mm")
        c.hline(x, x + w, by + head_h, dotted=False, width=c.px(2))
        lane_top = by + head_h
        lanes = lane_heights(chosen, block_h - head_h)
        for i in range(1, ncol):
            c.vline(x + gutter + col_w * i, lane_top, by + block_h)
        for k, (s, (ly, lh)) in enumerate(zip(chosen, lanes)):
            if k > 0:
                c.hline(x, x + w, lane_top + ly)
            if show_labels:
                name = c.fit(s["name"], label_font, gutter - c.px(6))
                c.d.text((x, lane_top + ly + c.px(4)), name, font=label_font, fill=0)
        for i, off in enumerate(offsets):
            day = wk[off]
            cx = x + gutter + col_w * i
            buckets = bucket_by_section(by_day.get(day, []), ids, default_id)
            for s, (ly, lh) in zip(chosen, lanes):
                entries = buckets.get(s["id"], [])
                if entries:
                    draw_lane_lines(c, cx + pad, lane_top + ly + pad, col_w - pad * 2, lh - pad,
                                    entries, cfg, size=12 if mode == "week" else 11,
                                    line_h=15 if mode == "week" else 13,
                                    wrap=2 if mode == "week" else 1)


# ------------------------------------------------------------ compact month

def draw_month_compact(c, x, y, w, h, today, by_day, cfg):
    first = first_weekday(cfg)
    title_h, wd_h = c.px(38), c.px(22)
    c.text((x, y), f"{pycal.month_name[today.month]} {today.year}", 26, bold=True)
    names = [pycal.day_abbr[(first + i) % 7] for i in range(7)]
    weeks = pycal.Calendar(first).monthdatescalendar(today.year, today.month)
    cell_w = w / 7.0
    cell_h = (h - title_h - wd_h) / len(weeks)
    gy = y + title_h
    for i, name in enumerate(names):
        c.text((x + cell_w * i + cell_w / 2, gy + wd_h - c.px(4)), name[0], 13, bold=True, anchor="ms")
    c.hline(x, x + w, gy + wd_h, dotted=False, width=c.px(2))
    top = gy + wd_h
    for r, week in enumerate(weeks):
        cy = top + cell_h * r
        for i, day in enumerate(week):
            if day.month != today.month:
                continue
            cx = x + cell_w * i
            n = len(by_day.get(day, []))
            mid_x, num_y = cx + cell_w / 2, cy + cell_h * 0.38
            is_today = day == today
            if is_today:
                rad = c.px(15)
                c.d.ellipse([mid_x - rad, num_y - rad, mid_x + rad, num_y + rad], fill=0)
            c.text((mid_x, num_y), str(day.day), 17, bold=True, fill=255 if is_today else 0, anchor="mm")
            dots, dot_r = min(n, 3), max(2, c.px(2.5))
            for k in range(dots):
                dx = mid_x + (k - (dots - 1) / 2.0) * c.px(9)
                dy = num_y + c.px(22)
                c.d.ellipse([dx - dot_r, dy - dot_r, dx + dot_r, dy + dot_r], fill=0)


# ------------------------------------------------------------------ agenda

def _day_label(day, today):
    if day == today:
        return "Today"
    if day == today + timedelta(days=1):
        return "Tomorrow"
    return f"{pycal.day_abbr[day.weekday()]} {day.day} {pycal.month_abbr[day.month]}"


def draw_agenda(c, x, y, w, h, today, by_day, cfg, sections, heading_right=None):
    names = {s["id"]: s["name"] for s in sections}
    default_id = sections[0]["id"]
    c.text((x, y), "Coming up", 26, bold=True)
    if heading_right:
        c.text((x + w, y + c.px(8)), heading_right, 14, bold=True, anchor="ra")
    y += c.px(38)
    c.hline(x, x + w, y, dotted=False, width=c.px(2))
    y += c.px(10)
    bottom = y - c.px(48) + h

    head_font, body_font, small_font = c.font(15, True), c.font(15), c.font(12)
    tf = cfg["time_format"]
    time_col = max(body_font.getlength(fmt_time("12:00", tf)), body_font.getlength("All day")) + c.px(18)
    row_h, loc_h, head_h = c.px(21), c.px(15), c.px(24)

    def meta(e):
        parts = []
        if e["section_id"] in names and e["section_id"] != default_id:
            parts.append(names[e["section_id"]])
        if e["location"]:
            parts.append(e["location"])
        return ", ".join(parts)

    rows = []
    for i in range(AGENDA_DAYS + 1):
        day = today + timedelta(days=i)
        entries = by_day.get(day, [])
        if entries:
            rows.append(("head", _day_label(day, today)))
            rows.extend(("event", e) for e in entries)

    if not rows:
        c.text((x, y + c.px(8)), "Nothing scheduled.", 16, bold=True)
        c.text((x, y + c.px(32)), f"The next {AGENDA_DAYS} days are clear.", 14)
        return

    def height(row):
        kind, val = row
        if kind == "head":
            return head_h
        return row_h + (loc_h if meta(val) else 0)

    more_h = c.px(22)
    used, cut = 0, len(rows)
    for i, row in enumerate(rows):
        remaining_events = sum(1 for r in rows[i:] if r[0] == "event")
        reserve = 0 if remaining_events == 0 else more_h
        if y + used + height(row) > bottom - (0 if i == len(rows) - 1 else reserve):
            cut = i
            break
        used += height(row)
    visible = rows[:cut]
    while visible and visible[-1][0] == "head":
        visible.pop()
    hidden = sum(1 for r in rows[len(visible):] if r[0] == "event")

    for kind, val in visible:
        if kind == "head":
            y += c.px(4)
            c.d.text((x, y), val, font=head_font, fill=0)
            y += head_h - c.px(4)
            continue
        label = "All day" if val["all_day"] else fmt_time(val["start_time"], tf)
        if val["all_day"] and val["day_count"] > 1:
            label = f"Day {val['day_number']}/{val['day_count']}"
        long_label = len(label) > 8
        c.d.text((x + c.px(2), y + (c.px(2) if long_label else 0)), label,
                 font=small_font if long_label else body_font, fill=0)
        c.d.text((x + time_col, y), c.fit(val["title"], body_font, w - time_col), font=body_font, fill=0)
        y += row_h
        extra = meta(val)
        if extra:
            c.d.text((x + time_col, y - c.px(2)), c.fit(extra, small_font, w - time_col), font=small_font, fill=0)
            y += loc_h

    if hidden:
        c.d.rectangle([x, bottom - more_h + c.px(4), x + w, bottom], fill=0)
        c.text((x + w / 2, bottom - more_h / 2 + c.px(2)), f"+{hidden} more", 13, bold=True, fill=255, anchor="mm")


# ------------------------------------------------------------------- views

def render_png(view, w, h, now, cfg=None, sections=None):
    cfg = cfg or store.get_settings()
    sections = sections or store.list_sections()
    view = view if view in VIEWS else cfg["view"]
    today = now.date()
    c = Canvas(w, h)
    m = c.px(22)

    first_of_month = today.replace(day=1)
    win_start = min(first_of_month - timedelta(days=14), today - timedelta(days=7))
    win_end = max(first_of_month + timedelta(days=45), today + timedelta(days=AGENDA_DAYS))
    by_day = store.events_by_day(win_start, win_end)

    stamp = None
    if cfg["show_updated"]:
        t = now.strftime("%H:%M") if cfg["time_format"] == "24" else now.strftime("%-I:%M %p")
        stamp = f"Updated {t}"
    foot = c.px(14) if stamp else 0

    if view in ("week", "twoweek", "month"):
        draw_grid(c, m, m, w - 2 * m, h - 2 * m - foot, view, today, by_day, sections, cfg)
    elif view == "agenda":
        draw_agenda(c, m, m, w - 2 * m, h - 2 * m - c.px(14), today, by_day, cfg, sections,
                    heading_right=f"{pycal.day_name[today.weekday()]} {today.day} {pycal.month_name[today.month]}")
    else:
        left_w = int(w * 0.43)
        draw_month_compact(c, m, m, left_w - m, h - 2 * m, today, by_day, cfg)
        div_x = left_w + c.px(10)
        c.vline(div_x, m, h - m)
        ax = div_x + c.px(22)
        draw_agenda(c, ax, m, w - ax - m, h - 2 * m - c.px(14), today, by_day, cfg, sections)
    if stamp:
        c.text((w - m, h - m + c.px(4)), stamp, 11, anchor="rs")
    return c.png()
