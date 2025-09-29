# PPT 번역 파이프라인 계획 (LangChain O3)

## 1. 목표
- Streamlit Cloud 단일 환경에서 **용어 전처리 → 번역 → 삽입** 전 과정을 처리한다.
- 번역 속도·비용을 최적화하면서 게임/마케팅 도메인 용어 일관성을 확보한다.
- 오버엔지니어링을 피하고 pure-python 의존성만 사용한다.

---

## 2. 환경 제약
| 항목 | 제약 |
|------|------|
| 실행 플랫폼 | Streamlit Cloud (CPU 1 core, RAM≈1 GB) |
| 외부 DB | 사용 불가 (내부 SQLite/파일 캐시만) |
| 대용량 모델 다운로드 | 불가 (>200 MB) |
| 컴파일·C++ 의존 패키지 | 설치 실패 위험 → 미사용 |

---

## 3. 전체 아키텍처
```mermaid
graph TD
  A[슬라이드 텍스트 추출] --> B[용어 후보 추출 (n-gram)]
  B --> C[OpenAI 용어 Mapping 1회]
  C --> D[자동 Glossary(JSON)]
  D --> E[LangChain ChatOpenAI 번역(batch)]
  E --> F[PPT에 번역 삽입]
```

### 주요 컴포넌트
1. **TermExtractor**  `regex n-gram + 빈도` → Top-N 용어 목록
2. **GlossaryBuilder**  `ChatOpenAI` JSON 응답으로 한–영 매핑 생성
3. **Translator**  `LangChain ChatOpenAI` + `JsonOutputParser`
4. **Cache**  `st.cache_data` (hash ⟶ glossary)

---

## 4. 단계별 상세
### 4.1 용어 후보 추출 (`preprocess_terms.py`)
```python
import re, collections

def extract_top_terms(texts: list[str], max_terms=100):
    tokens = re.findall(r"[A-Za-z가-힣]{2,}", " ".join(texts))
    noun_like = [t for t in tokens if not t.isdigit()]
    top = collections.Counter(noun_like).most_common(max_terms)
    return [w for w,_ in top]
```

### 4.2 Glossary 자동 작성
```python
from openai import OpenAI

def build_glossary(terms: list[str], target_lang="en") -> dict[str,str]:
    prompt = (
        "다음 용어를 영어로 번역하거나 브랜드면 원어 유지 후 JSON만 반환:"\n + ", ".join(terms)
    )
    client = OpenAI()
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role":"user","content":prompt}],
        response_format={"type":"json_object"}
    )
    return json.loads(resp.choices[0].message.content)["glossary"]
```

### 4.3 번역 체인 (LangChain)
```python
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import JsonOutputParser

llm = ChatOpenAI(model="gpt-4o-mini", temperature=0.1)
parser = JsonOutputParser()

SYSTEM = "You are a professional game-marketing PPT translator."
USER = (
    "Glossary: {glossary_json}\n"
    "Source hint: {src_hint}\nTarget lang: {tgt_lang}\n"
    "SOURCE(JSON): {source_json}"
)

prompt = ChatPromptTemplate.from_messages([("system", SYSTEM), ("user", USER)])
chain = prompt | llm.bind(response_format={"type":"json_object"}) | parser
```
사용:
```python
translated = chain.invoke({
    "glossary_json": json.dumps(glossary, ensure_ascii=False),
    "src_hint": "auto",
    "tgt_lang": "en",
    "source_json": json.dumps(batch, ensure_ascii=False),
})["result"]
```

---

## 5. 의존성 (requirements.txt 추가)
```txt
openai>=1.13
langchain>=0.1.19
langchain-openai>=0.0.5
python-dotenv
streamlit
```
(모두 pure-python wheel 제공)  
*추가 패키지 없음: spaCy / faiss / sklearn 제거*

---

## 6. TDD 목록
| 테스트 | 설명 |
|--------|------|
| `test_extract_terms.py` | 3개 언어 문단 ↔ 추출 결과 길이·중복 체크 |
| `test_build_glossary.py` | Mock LLM → JSON 구조·키 수 일치 확인 |
| `test_translate_chain.py` | 단일 batch 번역 결과 길이·글로서리 일관성 |

---

## 7. 작업 일정
| Day | Task |
|-----|------|
| 0.5 | `preprocess_terms.py` 작성 & 단위테스트 |
| 0.5 | `translate_lc.py` 체인 구현 |
| 0.5 | `ppt_generator.py` 호출부 대체, Streamlit 캐시 연동 |
| 0.5 | E2E 테스트 3종 & README 업데이트 |

총 **2일** 내 MVP 완성 가능.

---

## 8. 검증 기준
- 한글→영어 PPT(4000단어) 처리 시간 **≤ 60초** (gpt-4o-mini 기준)
- 동일 문장 재번역 비용 **0** (캐시 적중)
- 용어 일관성: 수동 샘플 50문장 중 오류 ≤ 1건

---

## 9. 추후 개선 TODO (선택)
- 단계적 QC 후처리(길이·형식만) LangGraph 도입(선택)
- 용어 목록을 Google Sheet로 관리 & 주기적 동기화
- OpenAI Function Calling으로 번역 결과를 PPT json schema에 바로 매핑

---

> 작성: o3-AI assistant · 2025-09-29
