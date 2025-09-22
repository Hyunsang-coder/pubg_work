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
    # 양방향 번역을 위한 개선된 프롬프트
    if config.target_lang == "auto":
        translate_instruction = "Translate each item: Korean to English, English to Korean"
    else:
        translate_instruction = f"Translate each item to {config.target_lang}"
    
    prompt = (
        f"{translate_instruction}. Keep numbers/dates/URLs/code unchanged. "
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
    # 양방향 번역을 위한 개선된 프롬프트
    if config.target_lang == "auto":
        translate_instruction = "Translate the following Markdown: Korean to English, English to Korean"
    else:
        translate_instruction = f"Translate the following Markdown into {config.target_lang}"
    
    prompt = (
        f"You are a professional translator. {translate_instruction}. "
        "Keep structure and punctuation. Do not change numbers, dates, URLs, or code blocks.\n"
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


# 이전 함수들은 더 이상 사용되지 않아 삭제됨 (reinsert_v2.py에서 직접 처리)
