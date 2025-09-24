from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from openai import OpenAI
from dotenv import load_dotenv, find_dotenv
LANGUAGE_NAMES = {
    "ko": "Korean",
    "en": "English",
    "ja": "Japanese",
    "zh": "Chinese",
}


def _language_name(code: Optional[str]) -> str:
    if not code:
        return ""
    return LANGUAGE_NAMES.get(code.lower(), code)


@dataclass
class TranslationConfig:
    source_lang: str = "auto"
    target_lang: str = "en"
    model: str = os.getenv("OPENAI_TRANSLATE_MODEL", "gpt-4o-mini")
    temperature: float = 0.1  # Keep low variance so JSON schema stays stable
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


MAX_TRANSLATION_RETRIES = 1


def _decode_translation_payload(content: str) -> List[str]:
    try:
        data = json.loads(content)
    except json.JSONDecodeError as exc:
        raise ValueError("Invalid JSON payload") from exc

    if not isinstance(data, dict):
        raise ValueError("Payload root is not an object")

    result: Any = data.get("result")
    if not isinstance(result, list):
        raise ValueError("Missing 'result' array")

    normalized: List[str] = []
    for item in result:
        if isinstance(item, str):
            normalized.append(item)
        elif item is None:
            normalized.append("")
        else:
            normalized.append(str(item))

    return normalized


def translate_texts(texts: List[str], config: TranslationConfig) -> List[str]:
    if not texts:
        return []
    client = _get_openai_client()
    use_temperature = config.model.lower() != "gpt-5" if config.model else True
    # In-batch duplicate caching: translate unique sentences once
    index_of: Dict[str, int] = {}
    unique_texts: List[str] = []
    back_refs: List[int] = []
    for t in texts:
        if t in index_of:
            back_refs.append(index_of[t])
        else:
            idx = len(unique_texts)
            index_of[t] = idx
            unique_texts.append(t)
            back_refs.append(idx)
    target_code = (config.target_lang or "en").lower()
    if target_code == "auto":
        target_name = "Auto (opposite language)"
    else:
        target_name = _language_name(target_code) or target_code

    source_code = (config.source_lang or "auto").lower()
    source_hint = ""

    if source_code == target_code and target_code != "auto":
        translate_instruction = f"Polish each item in {target_name} and improve clarity while keeping the meaning."
        source_hint = target_name
    elif source_code == "auto":
        translate_instruction = f"Detect whether each item is Korean, English, Japanese, or Chinese and translate it into {target_name}."
        source_hint = "Auto-detect (ko/en/ja/zh)"
    else:
        source_name = _language_name(source_code) or source_code
        translate_instruction = f"Translate each item from {source_name} to {target_name}."
        source_hint = source_name

    source_payload = unique_texts
    prompt = (
        f"{translate_instruction} Keep numbers/dates/URLs/code unchanged. "
        "Return only JSON with a 'result' array of translated strings matching the SOURCE order and length.\n"
        f"Source language hint: {source_hint}\n"
        f"Target language: {target_name}\n"
        f"Glossary(JSON): {json.dumps(config.glossary or {}, ensure_ascii=False)}\n"
        f"Extra Instructions: {config.extra_instructions or 'None'}\n"
        f"SOURCE(JSON): {json.dumps(source_payload, ensure_ascii=False)}"
    )

    messages = [
        {"role": "system", "content": "You are a translator. Output only valid JSON."},
        {"role": "user", "content": prompt},
    ]

    for attempt in range(MAX_TRANSLATION_RETRIES + 1):
        kwargs = {
            "model": config.model,
            "messages": messages,
            "response_format": {"type": "json_object"},
        }
        if use_temperature:
            kwargs["temperature"] = config.temperature
        resp = client.chat.completions.create(**kwargs)
        content = (resp.choices[0].message.content or "{}").strip()
        try:
            decoded = _decode_translation_payload(content)
            if len(decoded) != len(source_payload):
                raise ValueError("LENGTH_MISMATCH")
            return [decoded[i] for i in back_refs]
        except ValueError as exc:
            if exc.args and exc.args[0] == "LENGTH_MISMATCH":
                guidance = (
                    "번역 응답에 누락된 항목이 있어 실패했습니다. '번역된 PPT 생성'을 다시 실행하거나 다른 모델을 선택해주세요."
                )
                raise RuntimeError(guidance) from exc
            if attempt >= MAX_TRANSLATION_RETRIES:
                snippet = content[:2000]
                raise RuntimeError(f"번역 응답 파싱에 실패했습니다. 응답: {snippet}") from exc
            messages.append({
                "role": "system",
                "content": "Your previous reply was not valid JSON. Return only `{\"result\": [\"...\"]}` with the correct length."
            })
    raise RuntimeError("번역 응답 파싱에 실패했습니다.")



#deprecated
#def translate_markdown(md_text: str, config: TranslationConfig) -> str:
    client = _get_openai_client()
    glossary_text = json.dumps(config.glossary, ensure_ascii=False) if config.glossary else "{}"
    extra_text = f"Extra instructions: {config.extra_instructions}\n" if config.extra_instructions else ""
    target_code = (config.target_lang or "en").lower()
    if target_code == "auto":
        target_name = "Auto (opposite language)"
    else:
        target_name = _language_name(target_code) or target_code

    source_code = (config.source_lang or "auto").lower()
    source_hint = ""

    if source_code == target_code and target_code != "auto":
        translate_instruction = f"Polish the following Markdown in {target_name}, improving clarity while preserving meaning."
        source_hint = target_name
    elif source_code == "auto":
        translate_instruction = f"Detect whether each section is Korean, English, Japanese, or Chinese and translate the Markdown into {target_name}."
        source_hint = "Auto-detect (ko/en/ja/zh)"
    else:
        source_name = _language_name(source_code) or source_code
        translate_instruction = f"Translate the following Markdown from {source_name} into {target_name}."
        source_hint = source_name

    prompt = (
        f"You are a professional translator. {translate_instruction} "
        "Keep structure and punctuation. Do not change numbers, dates, URLs, or code blocks.\n"
        "Use the glossary strictly when applicable; do-not-translate terms must stay as-is.\n"
        f"Source language hint: {source_hint}\n"
        f"Target language: {target_name}\n"
        f"{extra_text}"
        f"Glossary(JSON): {glossary_text}\n\nSOURCE:\n```md\n{md_text}\n```"
    )
    messages = [
        {"role": "system", "content": "Professional translator. Output only translated Markdown."},
        {"role": "user", "content": prompt},
    ]
    kwargs = {"model": config.model, "messages": messages}
    if not (config.model and config.model.lower() == "gpt-5"):
        kwargs["temperature"] = config.temperature
    resp = client.chat.completions.create(**kwargs)
    return resp.choices[0].message.content or ""


# 이전 함수들은 더 이상 사용되지 않아 삭제됨 (reinsert_v2.py에서 직접 처리)
