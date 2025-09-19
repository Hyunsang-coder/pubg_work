from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Dict, List, Optional

from openai import OpenAI
from dotenv import load_dotenv, find_dotenv


@dataclass
class TranslationConfig:
    target_lang: str = "en"
    model: str = os.getenv("OPENAI_TRANSLATE_MODEL", "gpt-4o-mini")
    temperature: float = 0.0
    glossary: Optional[Dict[str, str]] = None
    extra_instructions: Optional[str] = None


def _get_openai_client() -> OpenAI:
    # Ensure .env is loaded from project root or parents
    load_dotenv(find_dotenv(usecwd=True), override=False)

    api_key = (
        os.getenv("OPENAI_API_KEY")
        or os.getenv("OPEN_API_KEY")  # common typo fallback
        or os.getenv("OPEN_AI_KEY")   # common typo fallback
    )
    if not api_key:
        # Try Streamlit secrets if available
        try:
            import streamlit as st  # type: ignore

            api_key = st.secrets.get("OPENAI_API_KEY")  # type: ignore
        except Exception:
            api_key = None

    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY is not set. Create a .env file (copy from .env.example) or set it in Streamlit secrets."
        )
    return OpenAI(api_key=api_key)


def build_prompt(markdown: str, glossary: Optional[Dict[str, str]], target_lang: str, extra: Optional[str] = None) -> str:
    glossary_text = json.dumps(glossary, ensure_ascii=False) if glossary else "{}"
    extra_text = f"Extra instructions: {extra}\n" if extra else ""
    return (
        "You are a professional translator. Translate the following Markdown into "
        f"{target_lang}. Keep structure and punctuation. Do not change numbers, dates, URLs, or code blocks.\n"
        "Use the glossary strictly when applicable; do-not-translate terms must stay as-is.\n"
        f"{extra_text}"
        f"Glossary(JSON): {glossary_text}\n\nSOURCE:\n```md\n{markdown}\n```"
    )


def translate_markdown(md_text: str, config: TranslationConfig) -> str:
    client = _get_openai_client()
    messages = [
        {"role": "system", "content": "Professional translator. Output only translated Markdown."},
        {"role": "user", "content": build_prompt(md_text, config.glossary, config.target_lang, config.extra_instructions)},
    ]
    resp = client.chat.completions.create(
        model=config.model,
        temperature=config.temperature,
        messages=messages,
    )
    return resp.choices[0].message.content or ""


def translate_texts(texts: List[str], config: TranslationConfig) -> List[str]:
    if not texts:
        return []
    client = _get_openai_client()
    prompt = (
        "Translate each item in the provided JSON list to "
        f"{config.target_lang}. Keep numbers/dates/URLs/code unchanged. "
        "Return ONLY a JSON array of strings with the same length/order.\n"
        f"Glossary(JSON): {json.dumps(config.glossary or {}, ensure_ascii=False)}\n"
        f"SOURCE(JSON): {json.dumps(texts, ensure_ascii=False)}"
    )
    resp = client.chat.completions.create(
        model=config.model,
        temperature=config.temperature,
        messages=[
            {"role": "system", "content": "Professional translator. Output only a valid JSON array."},
            {"role": "user", "content": prompt},
        ],
    )
    content = resp.choices[0].message.content or "[]"
    try:
        data = json.loads(content)
        if isinstance(data, list) and all(isinstance(x, str) for x in data):
            return data
    except Exception:
        pass
    lines = [l.strip() for l in content.splitlines() if l.strip()]
    if len(lines) == len(texts):
        return lines
    return texts


def shorten_line(line: str, max_chars: int, model: str | None = None) -> str:
    if len(line) <= max_chars:
        return line
    client = _get_openai_client()
    m = model or os.getenv("OPENAI_SHORTEN_MODEL", "gpt-4o-mini")
    prompt = (
        "Rephrase the sentence in English under "
        f"{max_chars} characters while preserving key meaning. Do not change numbers or URLs.\n"
        f"SOURCE: {line}"
    )
    resp = client.chat.completions.create(
        model=m,
        temperature=0,
        messages=[
            {"role": "system", "content": "You are a concise rephraser."},
            {"role": "user", "content": prompt},
        ],
    )
    out = (resp.choices[0].message.content or line).strip()
    return out[:max_chars] if len(out) > max_chars else out
