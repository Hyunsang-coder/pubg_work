from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Dict, List, Optional

from dotenv import find_dotenv, load_dotenv

os.environ.setdefault("OPENAI_USE_SIGNAL_TIMEOUT", "0")
_DOTENV_LOADED = load_dotenv(find_dotenv(usecwd=True), override=False)

from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import JsonOutputParser
from langchain_openai import ChatOpenAI
from pydantic import BaseModel, Field

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
    temperature: float = 0.1
    glossary: Optional[Dict[str, str]] = None
    extra_instructions: Optional[str] = None


class _TranslationPayload(BaseModel):
    result: List[str] = Field(..., description="Translated strings in original order")


def _build_translation_chain(config: TranslationConfig):
    parser = JsonOutputParser(pydantic_object=_TranslationPayload)
    prompt = ChatPromptTemplate.from_messages([
        (
            "system",
            "You are a meticulous translator for professional game development and marketing decks. "
            "Always respond with JSON only.\n\n{format_instructions}",
        ),
        (
            "user",
            "{instruction}\n"
            "Source language hint: {source_hint}\n"
            "Target language: {target_name}\n"
            "Glossary JSON: {glossary_json}\n"
            "Extra instructions: {extra_instructions}\n"
            "SOURCE(JSON): {source_json}",
        ),
    ]).partial(format_instructions=parser.get_format_instructions())

    temperature = config.temperature
    if config.model and config.model.lower() == "gpt-5":
        temperature = 0.0

    chat = ChatOpenAI(
        model=config.model,
        temperature=temperature,
        response_format={"type": "json_object"},
    )
    return prompt | chat | parser


def translate_texts(texts: List[str], config: TranslationConfig) -> List[str]:
    if not texts:
        return []

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
        translate_instruction = (
            f"Detect whether each item is Korean, English, Japanese, or Chinese and translate it into {target_name}."
        )
        source_hint = "Auto-detect (ko/en/ja/zh)"
    else:
        source_name = _language_name(source_code) or source_code
        translate_instruction = f"Translate each item from {source_name} to {target_name}."
        source_hint = source_name

    source_payload = unique_texts
    chain = _build_translation_chain(config)
    payload = {
        "instruction": translate_instruction + " Keep numbers/dates/URLs/code unchanged.",
        "source_hint": source_hint,
        "target_name": target_name,
        "glossary_json": json.dumps(config.glossary or {}, ensure_ascii=False),
        "extra_instructions": config.extra_instructions or "None",
        "source_json": json.dumps(source_payload, ensure_ascii=False),
    }

    try:
        raw_response = chain.invoke(payload)
    except Exception as exc:
        raise RuntimeError("번역 응답을 가져오는데 실패했습니다.") from exc

    if isinstance(raw_response, _TranslationPayload):
        response = raw_response
    elif isinstance(raw_response, dict):
        try:
            response = _TranslationPayload(**raw_response)
        except Exception as exc:
            raise RuntimeError("번역 응답 파싱에 실패했습니다.") from exc
    else:
        raise RuntimeError("예상치 못한 번역 응답 형식입니다.")

    decoded = response.result
    if len(decoded) != len(source_payload):
        raise RuntimeError("번역 응답 길이가 원본과 일치하지 않습니다.")

    return [decoded[i] for i in back_refs]
