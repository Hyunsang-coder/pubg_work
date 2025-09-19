from __future__ import annotations

from dataclasses import dataclass
from typing import List

from pptx import Presentation
from pptx.util import Pt

from .translate import shorten_line


@dataclass
class OverflowPolicy:
    # Try rephrase first (shorten) then fallback to font size reduce
    max_chars_per_paragraph: int = 180
    min_font_size_pt: int = 12
    reduce_step_pt: int = 1


def _reduce_font_size(text_frame, min_pt: int, step_pt: int):
    for p in text_frame.paragraphs:
        for run in p.runs:
            size = run.font.size.pt if run.font.size is not None else 18
            new_size = max(min_pt, int(size) - step_pt)
            run.font.size = Pt(new_size)


def _apply_text_to_shape(shape, text_lines: List[str], policy: OverflowPolicy):
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        return

    # 1) Rephrase long lines first
    adjusted: List[str] = []
    for line in text_lines:
        line = line or ""
        if len(line) > policy.max_chars_per_paragraph:
            line = shorten_line(line, policy.max_chars_per_paragraph)
        adjusted.append(line)

    tf = shape.text_frame
    tf.clear()

    # 2) Write text back
    for i, line in enumerate(adjusted):
        if i == 0:
            p = tf.paragraphs[0]
            p.text = line
        else:
            p = tf.add_paragraph()
            p.text = line

    # 3) If still long, reduce font size minimally
    if any(len(l) > policy.max_chars_per_paragraph for l in adjusted):
        _reduce_font_size(tf, policy.min_font_size_pt, policy.reduce_step_pt)


def create_translated_copy(input_pptx: str, translated_docs, output_pptx: str, policy: OverflowPolicy):
    prs = Presentation(input_pptx)
    for slide_idx, slide in enumerate(prs.slides):
        text_shapes = [s for s in slide.shapes if hasattr(s, "has_text_frame") and s.has_text_frame]
        doc = next((d for d in translated_docs if d.slide_index == slide_idx), None)
        if not doc:
            continue
        # flatten translated text lines in order
        new_lines: List[str] = []
        for b in doc.blocks:
            if getattr(b, "lines", None):
                new_lines.extend([l for l in b.lines if l is not None])
        line_idx = 0
        for shape in text_shapes:
            para_count = len(shape.text_frame.paragraphs)
            if para_count <= 0:
                continue
            assign = new_lines[line_idx: line_idx + para_count]
            if not assign:
                continue
            _apply_text_to_shape(shape, assign, policy)
            line_idx += len(assign)
    prs.save(output_pptx)
