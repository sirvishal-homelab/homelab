"""Domain objects. Pure data: no I/O, no knowledge of Beszel's JSON or the screen's API."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum


class SystemState(StrEnum):
    UP = "up"
    DOWN = "down"
    PAUSED = "paused"  # monitored-off on purpose: never a problem (owner, 2026-11-15)
    PENDING = "pending"


@dataclass(frozen=True)
class System:
    id: str
    name: str
    state: SystemState
    cpu_pct: float
    mem_pct: float


@dataclass(frozen=True)
class Container:
    system_id: str
    name: str  # key is system_id + name: Docker ids change when a container is recreated


@dataclass(frozen=True)
class Snapshot:
    systems: tuple[System, ...]
    containers: tuple[Container, ...]  # Beszel lists running containers only (docs/08 B-2)
    taken_at: datetime


class Level(StrEnum):
    OK = "ok"
    WARN = "warn"
    BAD = "bad"


class RowState(StrEnum):
    UP = "up"
    DOWN = "down"
    PAUSED = "paused"
    PENDING = "pending"
    MISSING = "missing"  # a focus system that Beszel does not report


@dataclass(frozen=True)
class Row:
    label: str
    state: RowState
    cpu_pct: float = 0.0
    mem_pct: float = 0.0
    cpu_level: Level = Level.OK
    mem_level: Level = Level.OK


@dataclass(frozen=True)
class Tile:
    up: int
    total: int
    paused: int = 0

    @property
    def level(self) -> Level:
        return Level.OK if self.up >= self.total else Level.BAD


@dataclass(frozen=True)
class Status:
    rows: tuple[Row, ...]
    containers: Tile
    others: Tile
    down_containers: tuple[str, ...] = field(default=())  # "system/name"
    down_systems: tuple[str, ...] = field(default=())  # labels; focus first, paused never
    updated_at: datetime | None = None
