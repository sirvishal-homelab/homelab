"""The scene a layout plans and the rasteriser draws (docs/04 section 5). Plain data, no drawing."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Rect:
    x: int
    y: int
    w: int
    h: int
    fill: int | None = None  # 0..255 grey, None = no fill
    outline: int | None = None


@dataclass(frozen=True)
class Line:
    x1: int
    y1: int
    x2: int
    y2: int
    gray: int = 0
    width: int = 1
    dotted: bool = False  # drawn pixel by pixel to stay crisp


@dataclass(frozen=True)
class Text:
    x: int
    y: int  # baseline
    text: str
    size: int = 12
    bold: bool = False
    gray: int = 0
    anchor: str = "l"  # "l", "m" or "r"


@dataclass(frozen=True)
class Glyph:
    """A small category symbol: square, triangle, circle, diamond, plus, cross, ring, box."""

    x: int
    y: int
    shape: str
    size: int = 8


@dataclass(frozen=True)
class Finding:
    code: str
    message: str
    where: str = ""


@dataclass
class FitReport:
    findings: list[Finding] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not any(f.code != "titles_truncated" for f in self.findings)

    def add(self, code: str, message: str, where: str = "") -> None:
        self.findings.append(Finding(code, message, where))


@dataclass
class Scene:
    width: int
    height: int
    items: list = field(default_factory=list)
    background: int = 255
