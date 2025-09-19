from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Dict, List, Optional

from openai import OpenAI
from dotenv import load_dotenv, find_dotenv
from .models import SlideDoc, TextBlock, TableBlock


@dataclass
class TranslationConfig:
    target_lang: str = "en"
    model: str = os.getenv("OPENAI_TRANSLATE_MODEL", "gpt-4o-mini")
    temperature: float = 1  # Some models don't support 0.0
    glossary: Optional[Dict[str, str]] = None
    extra_instructions: Optional[str] = None


def _get_openai_client() -> OpenAI:
    load_dotenv(find_dotenv(usecwd=True), override=False)
    api_key = os.getenv("OPENAI_API_KEY") or os.getenv("OPEN_API_KEY")
    if not api_key:
        try:
            import streamlit as st  # type: ignore
            api_key = st.secrets.get("OPENAI_API_KEY")
        except Exception:
            api_key = None
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not set in .env or Streamlit secrets.")
    return OpenAI(api_key=api_key)


def translate_texts(texts: List[str], config: TranslationConfig) -> List[str]:
    if not texts:
        return []
    client = _get_openai_client()
    prompt = (
        "Translate each item in the provided JSON list to "
        f"{config.target_lang}. Keep numbers/dates/URLs/code unchanged. "
        "Return ONLY a valid JSON object like `{\"result\": [\"...\", \"...\"]}` with the same length/order.\n"
        f"Glossary(JSON): {json.dumps(config.glossary or {}, ensure_ascii=False)}\n"
        f"Extra Instructions: {config.extra_instructions or 'None'}\n"
        f"SOURCE(JSON): {json.dumps(texts, ensure_ascii=False)}"
    )
    resp = client.chat.completions.create(
        model=config.model,
        temperature=config.temperature,
        messages=[{"role": "system", "content": "You are a translator. Output only a valid JSON object."}, {"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
    )
    content = resp.choices[0].message.content or '{"result": []}'
    try:
        data = json.loads(content)
        key = next((k for k in data if isinstance(data.get(k), list)), None)
        if key:
            result = data[key]
            if all(isinstance(x, str) for x in result) and len(result) == len(texts):
                return result
    except (json.JSONDecodeError, StopIteration):
        pass
    return texts  # Fallback


def translate_markdown(md_text: str, config: TranslationConfig) -> str:
    client = _get_openai_client()
    glossary_text = json.dumps(config.glossary, ensure_ascii=False) if config.glossary else "{}"
    extra_text = f"Extra instructions: {config.extra_instructions}\n" if config.extra_instructions else ""
    prompt = (
        "You are a professional translator. Translate the following Markdown into "
        f"{config.target_lang}. Keep structure and punctuation. Do not change numbers, dates, URLs, or code blocks.\n"
        "Use the glossary strictly when applicable; do-not-translate terms must stay as-is.\n"
        f"{extra_text}"
        f"Glossary(JSON): {glossary_text}\n\nSOURCE:\n```md\n{md_text}\n```"
    )
    messages = [
        {"role": "system", "content": "Professional translator. Output only translated Markdown."},
        {"role": "user", "content": prompt},
    ]
    resp = client.chat.completions.create(
        model=config.model,
        temperature=config.temperature,
        messages=messages,
    )
    return resp.choices[0].message.content or ""


def shorten_line(line: str, max_chars: int, model: str | None = None) -> str:
    try:
        if line is None:
            return ""
        text = str(line)
        if len(text) <= max_chars:
            return text
        client = _get_openai_client()
        m = model or os.getenv("OPENAI_SHORTEN_MODEL", "gpt-4o-mini")
        prompt = (
            "Rephrase the sentence in English under "
            f"{max_chars} characters while preserving key meaning. Do not change numbers or URLs.\n"
            f"SOURCE: {text}"
        )
        resp = client.chat.completions.create(
            model=m,
            temperature=0.1,
            messages=[
                {"role": "system", "content": "You are a concise rephraser."},
                {"role": "user", "content": prompt},
            ],
        )
        out = (resp.choices[0].message.content or text).strip()
        return out[:max_chars] if len(out) > max_chars else out
    except Exception:
        # Fallback to safe truncation
        try:
            s = "" if line is None else str(line)
            return s[:max_chars]
        except Exception:
            return ""


def orchestrate_translation(docs: List[SlideDoc], config: TranslationConfig) -> List[SlideDoc]:
    # 1. Flatten all text from docs into a single list
    texts_to_translate: List[str] = []
    source_map = []  # To map translations back
    for i, doc in enumerate(docs):
        for j, block in enumerate(doc.blocks):
            if isinstance(block, TextBlock):
                texts_to_translate.extend(block.lines)
                for k in range(len(block.lines)):
                    source_map.append((i, j, "text", k))
            elif isinstance(block, TableBlock):
                for r, row in enumerate(block.rows):
                    texts_to_translate.extend(row)
                    for c in range(len(row)):
                        source_map.append((i, j, "table", r, c))

    # 2. Translate in one batch
    translated_texts = translate_texts(texts_to_translate, config)

    # 3. Reconstruct translated docs using the source map
    from copy import deepcopy
    translated_docs: List[SlideDoc] = deepcopy(docs)
    text_idx = 0
    for i, doc in enumerate(translated_docs):
        for j, block in enumerate(doc.blocks):
            if isinstance(block, TextBlock):
                num_lines = len(block.lines)
                block.lines = translated_texts[text_idx : text_idx + num_lines]
                text_idx += num_lines
            elif isinstance(block, TableBlock):
                for r, row in enumerate(block.rows):
                    num_cells = len(row)
                    block.rows[r] = translated_texts[text_idx : text_idx + num_cells]
                    text_idx += num_cells

    return translated_docs
