from __future__ import annotations

from typing import List
from .models import SlideDoc, TextBlock, TableBlock, FigureBlock, NoteBlock


def _md_escape(text: str) -> str:
    # Simple escape for pipes in tables
    return text.replace("|", "\\|")


def blocks_to_markdown(blocks: List[object], options) -> str:
    lines: List[str] = []
    for block in blocks:
        if isinstance(block, TextBlock):
            for idx, line in enumerate(block.lines):
                line = line.strip()
                if not line:
                    continue
                level = 0
                if block.indent_levels and idx < len(block.indent_levels):
                    level = block.indent_levels[idx]
                lines.append(f"{'  ' * level}- {line}")
            lines.append("")
        elif isinstance(block, TableBlock):
            if not block.rows:
                continue
            header = block.rows[0] if block.has_header and block.rows else None
            body = block.rows[1:] if header else block.rows
            if header:
                lines.append("| " + " | ".join(_md_escape(c) for c in header) + " |")
                lines.append("| " + " | ".join(["---"] * len(header)) + " |")
            for row in body:
                lines.append("| " + " | ".join(_md_escape(c) for c in row) + " |")
            lines.append("")
        elif isinstance(block, FigureBlock):
            if block.figure_type == "image":
                if options.figures == "placeholder":
                    title = block.title or "Image"
                    lines.append(f"[Figure: {title}]")
            elif block.figure_type == "chart":
                if options.charts == "labels":
                    title = block.title or "Chart"
                    lines.append(f"[Figure: Chart, title=\"{title}\"]")
                elif options.figures == "placeholder":
                    lines.append("[Figure: Chart]")
            lines.append("")
        elif isinstance(block, NoteBlock):
            lines.append("> NOTE: " + block.text.replace("\n", " ").strip())
            lines.append("")
    return "\n".join(lines).strip() + "\n"


def docs_to_markdown(docs: List[SlideDoc], options) -> str:
    out: List[str] = []
    for doc in docs:
        out.append(f"## Slide {doc.slide_index + 1} - {doc.title}")
        out.append("")
        out.append(blocks_to_markdown(doc.blocks, options).rstrip())
        out.append("")
    return "\n".join(out).rstrip() + "\n"
