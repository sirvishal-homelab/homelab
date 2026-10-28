"""plan -> rasterise -> quantise -> PNG. The preview and the push both call this (docs/04 s2)."""

from __future__ import annotations

import hashlib
from collections import OrderedDict
from dataclasses import dataclass

from trmnl_calendar.render.layouts.base import PlanContext
from trmnl_calendar.render.layouts.registry import LAYOUTS
from trmnl_calendar.render.png import encode_gray_png
from trmnl_calendar.render.quantise import to_samples
from trmnl_calendar.render.raster import rasterise
from trmnl_calendar.render.scene import FitReport, Scene


@dataclass(frozen=True)
class RenderResult:
    png: bytes
    hash: str
    fit: FitReport
    scene: Scene
    width: int
    height: int
    bit_depth: int


def render(ctx: PlanContext, layout_id: str = "week_lanes") -> RenderResult:
    layout = LAYOUTS[layout_id]
    scene, fit = layout.plan(ctx)
    samples = to_samples(rasterise(scene), ctx.bit_depth)
    png = encode_gray_png(samples, scene.width, scene.height, ctx.bit_depth)
    digest = hashlib.sha256(png).hexdigest()[:24]
    return RenderResult(png, digest, fit, scene, scene.width, scene.height, ctx.bit_depth)


class RenderCache:
    """Small LRU of hash -> PNG so /render/<hash>.png never re-renders (docs/06 section 4)."""

    def __init__(self, size: int = 16):
        self._size = size
        self._items: OrderedDict[str, bytes] = OrderedDict()

    def put(self, key: str, png: bytes) -> None:
        self._items[key] = png
        self._items.move_to_end(key)
        while len(self._items) > self._size:
            self._items.popitem(last=False)

    def get(self, key: str) -> bytes | None:
        png = self._items.get(key)
        if png is not None:
            self._items.move_to_end(key)
        return png
