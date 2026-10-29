"""Boxed initials for the people on an occurrence, shared by the layouts (docs/04 section 3).

Every box is white with an outline (grey was hard to read). The main person comes first, then the
others, at most three.
"""

from __future__ import annotations

from trmnl_calendar.domain.models import Occurrence
from trmnl_calendar.render.layouts.base import PlanContext
from trmnl_calendar.render.scene import Rect, Text

BADGE_H = 12
MAX_BADGES = 3


def badge_spec(ctx: PlanContext, o: Occurrence) -> list[tuple[str, int, bool]]:
    """(label, width, is_owner) for the person badges of an occurrence, owner first."""
    if not ctx.opt("show_badges"):
        return []
    out = []
    for pid in o.people[:MAX_BADGES]:
        person = ctx.people.get(pid)
        if person:
            width = 12 if len(person.label) == 1 else 18
            out.append((person.label, width, pid == o.owner_id))
    return out


def badge_width(spec: list[tuple[str, int, bool]]) -> int:
    return sum(w for _, w, _ in spec) + max(0, len(spec) - 1)


def draw_badges(add, spec: list[tuple[str, int, bool]], right: int, y: int) -> None:
    """Boxed initials, right-aligned at x=right; all boxes are white; the main person is first."""
    x = right - badge_width(spec)
    for label, width, _is_owner in spec:
        add(Rect(x, y, width, BADGE_H, fill=255, outline=0))
        add(Text(x + width // 2, y + 10, label, 10, True, anchor="m"))
        x += width + 1
