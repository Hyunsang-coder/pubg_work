from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Literal


@dataclass
class ExtractOptions:
    with_notes: bool = False
    figures: Literal["placeholder", "omit"] = "placeholder"
    charts: Literal["labels", "placeholder", "omit"] = "labels"
    table_header: bool = True
    slide_range: Optional[List[int]] = None
