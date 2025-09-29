from __future__ import annotations

import json
import os
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence

from dotenv import find_dotenv, load_dotenv

os.environ.setdefault("OPENAI_USE_SIGNAL_TIMEOUT", "0")
_DOTENV_LOADED = load_dotenv(find_dotenv(usecwd=True), override=False)

from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import JsonOutputParser
from langchain_openai import ChatOpenAI
from pydantic import BaseModel, Field

from .models import SlideDoc, TextBlock, TableBlock, NoteBlock


@dataclass
class TermCandidate:
    term: str
    slide_index: int
    context: str
    score: int = 1


@dataclass
class PreflightTerm:
    source_term: str
    preferred_translation: Optional[str] = None
    category: Optional[str] = None
    rationale: Optional[str] = None


@dataclass
class PreflightResult:
    terms: List[PreflightTerm] = field(default_factory=list)
    style_note: str = ""
    ambiguous_spots: List[str] = field(default_factory=list)
    raw_candidates: List[TermCandidate] = field(default_factory=list)


class _PreflightLLMTerm(BaseModel):
    source_term: str = Field(..., description="The original terminology or phrase")
    preferred_translation: Optional[str] = Field(
        default=None,
        description="Recommended translation; leave null if unsure",
    )
    category: Optional[str] = Field(
        default=None,
        description="Category such as product, feature, marketing metric, event, etc.",
    )
    rationale: Optional[str] = Field(
        default=None,
        description="Brief reasoning for the recommendation",
    )


class _PreflightLLMResponse(BaseModel):
    terms: List[_PreflightLLMTerm] = Field(default_factory=list)
    style_note: str = Field(
        default="",
        description="Tone, register, and stylistic guidance for the translation phase",
    )
    ambiguous_spots: List[str] = Field(
        default_factory=list,
        description="Sentences or concepts that need clarification",
    )


_TERM_PATTERNS: Sequence[re.Pattern[str]] = (
    re.compile(r"\b[A-Z][A-Z0-9]{2,}\b"),  # Acronyms (AAA, KPI)
    re.compile(r"\b[A-Z][a-z]+(?: [A-Z][a-z]+)+\b"),  # Proper nouns with spaces (Battle Royale)
    re.compile(r"\b[A-Za-z]+(?:-[A-Za-z0-9]+)+\b"),  # Hyphenated terms (free-to-play)
    re.compile(r"\b[A-Za-z]*\d+[A-Za-z]*\b"),  # Tokens containing digits (Tier-1, GDC2025)
    re.compile(r"[가-힣]{4,}")  # Longer Korean words that may indicate specific concepts
)


def _iter_text_segments(docs: Iterable[SlideDoc]) -> Iterable[tuple[int, str, str]]:
    """Yield `(slide_index, slide_title, text)` for every textual element."""
    for doc in docs:
        title = doc.title
        for block in doc.blocks:
            if isinstance(block, TextBlock):
                for line in block.lines:
                    text = line.strip()
                    if text:
                        yield (doc.slide_index, title, text)
            elif isinstance(block, TableBlock):
                for row in block.rows:
                    text = " | ".join(cell.strip() for cell in row if cell)
                    if text:
                        yield (doc.slide_index, title, text)
            elif isinstance(block, NoteBlock):
                text = block.text.strip()
                if text:
                    yield (doc.slide_index, f"{title} (Note)", text)


def _extract_terms_from_text(text: str) -> List[str]:
    raw_terms = set()
    for pattern in _TERM_PATTERNS:
        matches = pattern.findall(text)
        for match in matches:
            cleaned = match.strip()
            if cleaned:
                raw_terms.add(cleaned)
    return list(raw_terms)


def collect_term_candidates(
    docs: Iterable[SlideDoc],
    *,
    max_terms: int = 50,
) -> List[TermCandidate]:
    """Collect candidate terminology with simple heuristics."""
    aggregates: Dict[str, Dict[str, object]] = defaultdict(lambda: {
        "count": 0,
        "contexts": [],
        "first_slide": None,
    })

    for slide_index, title, text in _iter_text_segments(docs):
        snippet = text if len(text) <= 160 else text[:157] + "..."
        for term in _extract_terms_from_text(text):
            key = term.lower()
            bucket = aggregates[key]
            bucket["count"] = int(bucket["count"]) + 1
            contexts: List[str] = bucket["contexts"]  # type: ignore[assignment]
            if len(contexts) < 3:
                contexts.append(f"Slide {slide_index + 1} — {snippet}")
            if bucket["first_slide"] is None:
                bucket["first_slide"] = slide_index
            bucket.setdefault("term", term)

    ranked = sorted(
        (
            TermCandidate(
                term=bucket.get("term", term_key),
                slide_index=int(bucket.get("first_slide") or 0),
                context="; ".join(bucket.get("contexts", [])),
                score=int(bucket.get("count", 0)),
            )
            for term_key, bucket in aggregates.items()
        ),
        key=lambda candidate: candidate.score,
        reverse=True,
    )
    return ranked[:max_terms]


def _build_outline(docs: Iterable[SlideDoc], *, line_limit: int = 40) -> str:
    lines: List[str] = []
    for doc in docs:
        header = f"Slide {doc.slide_index + 1}: {doc.title}"
        lines.append(header)
        appended = 0
        for block in doc.blocks:
            if isinstance(block, TextBlock):
                for line in block.lines:
                    text = line.strip()
                    if text:
                        lines.append(f"- {text}")
                        appended += 1
                        if appended >= 3:
                            break
            if appended >= 3:
                break
    return "\n".join(lines[:line_limit])


def _build_preflight_chain(model_name: Optional[str] = None):
    parser = JsonOutputParser(pydantic_object=_PreflightLLMResponse)
    prompt = ChatPromptTemplate.from_messages([
        (
            "system",
            "You are a bilingual localization strategist for game publishing decks. "
            "Your job is to review terminology candidates, recommend consistent translations, "
            "and highlight any ambiguous content. Always reply with JSON only.\n\n"
            "{format_instructions}",
        ),
        (
            "user",
            "Presentation outline:\n{outline}\n\n"
            "Terminology candidates (JSON):\n{candidates_json}\n\n"
            "Target translation language: {target_lang}.\n"
            "Return any important glossary entries with suggested translations that suit "
            "professional marketing/game development tone.",
        ),
    ]).partial(format_instructions=parser.get_format_instructions())

    chat = ChatOpenAI(
        model=model_name or os.getenv("OPENAI_PREFLIGHT_MODEL", "gpt-4o-mini"),
        temperature=0.2,
    )
    return prompt | chat | parser


def run_preflight(
    docs: Sequence[SlideDoc],
    *,
    target_lang: str = "en",
    model_name: Optional[str] = None,
    max_terms: int = 50,
) -> PreflightResult:
    """Run the terminology preflight step. Falls back gracefully on failure."""
    docs = list(docs)
    if not docs:
        return PreflightResult()

    candidates = collect_term_candidates(docs, max_terms=max_terms)
    outline = _build_outline(docs)

    if not candidates:
        return PreflightResult(raw_candidates=[])

    chain = _build_preflight_chain(model_name)
    payload = {
        "outline": outline,
        "candidates_json": json.dumps(
            [
                {
                    "term": candidate.term,
                    "score": candidate.score,
                    "context": candidate.context,
                    "slide": candidate.slide_index + 1,
                }
                for candidate in candidates
            ],
            ensure_ascii=False,
        ),
        "target_lang": target_lang,
    }

    try:
        raw_response = chain.invoke(payload)
    except Exception:
        return PreflightResult(raw_candidates=candidates)

    if isinstance(raw_response, _PreflightLLMResponse):
        response = raw_response
    elif isinstance(raw_response, dict):
        try:
            response = _PreflightLLMResponse(**raw_response)
        except Exception:
            return PreflightResult(raw_candidates=candidates)
    else:
        return PreflightResult(raw_candidates=candidates)

    terms = []
    for item in response.terms:
        source = (item.source_term or "").strip()
        if not source:
            continue
        terms.append(
            PreflightTerm(
                source_term=source,
                preferred_translation=(item.preferred_translation or None),
                category=(item.category or None),
                rationale=(item.rationale or None),
            )
        )

    return PreflightResult(
        terms=terms,
        style_note=response.style_note.strip(),
        ambiguous_spots=[spot.strip() for spot in response.ambiguous_spots if spot.strip()],
        raw_candidates=candidates,
    )


__all__ = [
    "TermCandidate",
    "PreflightTerm",
    "PreflightResult",
    "collect_term_candidates",
    "run_preflight",
]
