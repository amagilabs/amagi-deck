"""Plain data classes for a parsed deck. No I/O here."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import List, Optional


@dataclass
class Series:
    name: str
    values: List[Optional[float]]
    x_values: List[Optional[float]] = field(default_factory=list)


@dataclass
class Chart:
    kind: str
    title: str = ""
    categories: List[str] = field(default_factory=list)
    series: List[Series] = field(default_factory=list)


@dataclass
class Table:
    rows: List[List[str]]


@dataclass
class Picture:
    name: str
    alt_text: str = ""
    content_type: str = ""


@dataclass
class TextBlock:
    """One shape's text. Each entry in ``lines`` is (indent level, text)."""

    shape: str
    lines: List[tuple]


@dataclass
class Slide:
    number: int
    title: str = ""
    hidden: bool = False
    layout: str = ""
    text: List[TextBlock] = field(default_factory=list)
    notes: str = ""
    tables: List[Table] = field(default_factory=list)
    charts: List[Chart] = field(default_factory=list)
    pictures: List[Picture] = field(default_factory=list)
    links: List[str] = field(default_factory=list)
    image: str = ""


@dataclass
class Deck:
    source: str
    kind: str
    slides: List[Slide]
    width_in: float = 0.0
    height_in: float = 0.0
    title: str = ""
    renderer: str = ""
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)
