"""Pure rules (docs/05 section 4): snapshot + settings + container memory -> Status."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from esp_screen.domain.model import (
    Container,
    Level,
    Row,
    RowState,
    Snapshot,
    Status,
    System,
    SystemState,
    Tile,
)

MAX_FOCUS = 4


@dataclass(frozen=True)
class Thresholds:
    warn_pct: float = 75
    crit_pct: float = 90


DEFAULT_THRESHOLDS = Thresholds()


def level(pct: float, t: Thresholds) -> Level:
    if pct >= t.crit_pct:
        return Level.BAD
    if pct >= t.warn_pct:
        return Level.WARN
    return Level.OK


def find_focus(systems: Sequence[System], wanted: Sequence[str]) -> list[System | str]:
    """Match focus entries to systems by id, then by name (case-insensitive). Unknown -> str."""
    by_id = {s.id: s for s in systems}
    by_name = {s.name.lower(): s for s in systems}
    out: list[System | str] = []
    for w in list(wanted)[:MAX_FOCUS]:
        out.append(by_id.get(w) or by_name.get(w.lower()) or w)
    return out


def row_for(item: System | str, t: Thresholds, label: str | None = None) -> Row:
    if isinstance(item, str):
        return Row(label or item, RowState.MISSING)
    state = RowState(item.state.value)
    if state is not RowState.UP:
        return Row(label or item.name, state)
    return Row(
        label or item.name,
        RowState.UP,
        item.cpu_pct,
        item.mem_pct,
        level(item.cpu_pct, t),
        level(item.mem_pct, t),
    )


@dataclass(frozen=True)
class ContainerMemory:
    """What the app remembers per container: consecutive reads it has been missing."""

    missing: Mapping[tuple[str, str], int]
    allowed_to_stop: frozenset[tuple[str, str]] = frozenset()


def remember(
    memory: ContainerMemory, running: Sequence[Container], down_systems: set[str]
) -> dict[tuple[str, str], int]:
    """Next missing-count per known container. New containers are learned on sight.

    A container on a system that is not up is left unchanged: its absence says nothing.
    """
    seen = {(c.system_id, c.name) for c in running}
    nxt: dict[tuple[str, str], int] = {}
    for key, count in memory.missing.items():
        if key in seen:
            nxt[key] = 0
        elif key[0] in down_systems:
            nxt[key] = count
        else:
            nxt[key] = count + 1
    for key in seen:
        nxt.setdefault(key, 0)
    return nxt


def evaluate(
    snap: Snapshot,
    focus: Sequence[str],
    memory: Mapping[tuple[str, str], int],
    *,
    allowed_to_stop: frozenset[tuple[str, str]] = frozenset(),
    missing_limit: int = 2,
    thresholds: Thresholds = DEFAULT_THRESHOLDS,
    labels: Mapping[str, str] | None = None,
) -> Status:
    """`memory` is the result of `remember()` for this snapshot."""
    labels = labels or {}
    focus_items = find_focus(snap.systems, focus)
    focus_ids = {i.id for i in focus_items if isinstance(i, System)}
    rows = tuple(
        row_for(i, thresholds, labels.get(i.id) if isinstance(i, System) else None)
        for i in focus_items
    )

    others = [s for s in snap.systems if s.id not in focus_ids]
    paused = sum(1 for s in others if s.state is SystemState.PAUSED)
    watched = [s for s in others if s.state is not SystemState.PAUSED]
    other_tile = Tile(
        up=sum(1 for s in watched if s.state is SystemState.UP), total=len(watched), paused=paused
    )

    names = {s.id: s.name for s in snap.systems}
    expected = {k for k in memory if k not in allowed_to_stop}
    down = sorted(k for k in expected if memory[k] >= missing_limit)
    # A container counts as down only once it has been missing for `missing_limit` reads,
    # so one slow read or a quick restart does not flash the tile red.
    container_tile = Tile(up=len(expected) - len(down), total=len(expected))
    return Status(
        rows=rows,
        containers=container_tile,
        others=other_tile,
        down_containers=tuple(f"{names.get(sid, sid)}/{name}" for sid, name in down),
        updated_at=snap.taken_at,
    )
