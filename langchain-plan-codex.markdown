# LangChain 전처리/번역 개선 계획

## 1. 목표
- PPT 번역 파이프라인에 전처리 체인을 도입해 용어 일관성과 번역 품질을 향상한다.
- LangChain 기반 체인을 사용하되 Streamlit Cloud 환경(단일 인스턴스, 제한된 패키지 설치 시간)에서 안정적으로 동작하도록 단순한 구조로 설계한다.
- 기존 번역 흐름을 크게 변경하지 않고 `TranslationConfig`에 Glossary와 스타일 지침을 주입하는 방식으로 통합한다.

## 2. 구현 범위
1. **전처리 체인 구성**
   - `src/pptx2md/preflight.py` 신규 모듈 생성.
   - LangChain Runnable 시퀀스를 활용해 “규칙 기반 후보 수집 → LLM 정제 → Glossary 제안”을 수행.
   - 결과는 `PreflightResult`(`terms`, `style_note`, `ambiguous_spots`) 데이터 클래스로 반환.
2. **Streamlit 연동**
   - UI에 "용어 분석" 버튼과 결과 미리보기(표/텍스트)를 추가.
   - 사용자가 Glossary 제안을 편집하거나 승인할 수 있게 하며, 승인 결과를 세션 상태에 저장.
3. **번역 단계 통합**
   - `TranslationConfig.glossary`와 `extra_instructions`를 전처리 결과로 업데이트.
   - Glossary가 존재할 때만 LangChain 번역 체인을 호출하고, 나머지 로직은 기존과 동일하게 유지.
4. **검증 및 로그**
   - 소형 PPT(예: 3~5 슬라이드)로 수동 검증.
   - Glossary 적용 여부와 LLM 응답을 Streamlit sidebar 또는 로그로 표시.

## 3. LangChain 구성 요소
- **의존성**: `langchain`, `langchain-openai` (필수), `pydantic`(이미 사용 중인지 확인 후 필요 시 버전 고정).
- **전처리 체인 (RunnableSequence)**
  1. **CollectCandidates (RunnableLambda)**
     - `extract_pptx_to_docs` 출력에서 텍스트 블록을 순회.
     - 정규식 및 간단한 규칙(대문자 연속, 숫자+단위, 고유명사 패턴, 한글 명사 2~3어절)으로 후보 용어 리스트 생성.
     - 슬라이드 인덱스, 문장 컨텍스트를 함께 반환.
  2. **Aggregate (RunnableLambda)**
     - 후보를 소문자/공백 정규화하여 중복 제거.
     - 각 용어별 대표 원형, 출현 슬라이드, 예시 문장을 정리.
  3. **LLMRefine (ChatPromptTemplate + ChatOpenAI)**
     - Prompt: “게임 개발·마케팅 PPT 맥락에 맞는 번역 용어 추천, 톤&스타일 요약”.
     - Input: 후보 리스트(상위 N개), 내용 요약(슬라이드별 제목/핵심 문장).
     - Output: JSON(`terms`, `style_note`, `ambiguous_spots`).
     - `JsonOutputParser`로 구조 검증.

- **번역 체인 (재사용)**
  - 기존 `translate_texts` 함수 내부에 LangChain Translator 체인(`prompt | ChatOpenAI | JsonOutputParser`) 도입.
  - 전처리 결과 Glossary/스타일 지침을 prompt에 삽입.

## 4. Streamlit 흐름 변경
1. 업로드한 PPT를 세션 상태에 저장(기존 기능 유지).
2. “용어 분석” 클릭 시:
   - 전처리 체인 실행 → 결과를 `st.session_state.preflight_result`에 저장.
   - Glossary 표(st.dataframe)와 스타일 요약(st.text_area) 제공.
   - 사용자가 수정 후 "적용" 클릭 시 확정 Glossary/스타일을 세션에 반영.
3. 번역 실행 시:
   - 세션에 저장된 Glossary/스타일이 있으면 `TranslationConfig`에 주입.
   - 없으면 기존 흐름으로 진행.

## 5. 품질 및 성능 고려
- **Streamlit Cloud 제약 대응**
  - LangGraph 등 추가 의존성 미도입, Runnable 시퀀스만 사용.
  - LLM 호출 횟수 최소화: 전처리 LLM은 프레젠테이션당 1회.
  - 후보 리스트는 상위 40~50개로 제한해 prompt 길이를 관리.
- **에러 대응**
  - LLM JSON 파싱 실패 시 기본 Glossary 없음 상태로 폴백하고 사용자에게 경고.
  - Glossary 편집 중 예외 발생 시 저장하지 않고 다시 입력하도록 안내.
- **일관성**
  - Glossary는 딕셔너리로 저장하며, 번역 단계에서 키 그대로 prompt에 전달.
  - 스타일 노트는 translator prompt 상단에 삽입.

## 6. 작업 순서
1. `requirements.txt` 업데이트 (`langchain`, `langchain-openai`).
2. `src/pptx2md/preflight.py` 작성 및 단위 테스트(간단한 텍스트 입력 기반).
3. `translate_texts`에 LangChain Translator 체인 통합.
4. Streamlit UI 수정: 전처리 실행 및 결과 편집 기능 추가.
5. 로컬/Cloud에서 소형 PPT로 전체 흐름 검증.
6. README에 “전처리 → 번역” 사용법 및 Glossary 편집 안내 추가.

## 7. 후속 과제(선택)
- Glossary 캐시: 빈번한 용어를 로컬 JSON에 저장해 재사용.
- Terminology QA: 번역 후 Glossary 미준수 항목을 표시하는 간단한 체크(추후).
- 자동 테스트: 샘플 PPT를 fixtures로 두고 CLI 테스트 스크립트 작성.
