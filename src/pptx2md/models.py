from dataclasses import dataclass, field
from typing import List, Literal, Optional

BlockKind = Literal["text", "table", "figure", "note"]

@dataclass
class TextBlock:
    shape_id: str
    lines: List[str]
    indent_levels: List[int] = field(default_factory=list)

@dataclass
class TableBlock:
    shape_id: str
    rows: List[List[str]]
    has_header: bool = True

@dataclass
class FigureBlock:
    shape_id: str
    figure_type: Literal["image", "chart"]
    title: Optional[str] = None

@dataclass
class NoteBlock:
    text: str

@dataclass
class SlideDoc:
    slide_index: int
    title: str
    blocks: List[object] = field(default_factory=list)
