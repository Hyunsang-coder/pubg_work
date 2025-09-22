from __future__ import annotations
import os, time, json, io, uuid, hashlib, html
from datetime import datetime
import streamlit as st
from dotenv import load_dotenv
import sys
try:
    import pandas as pd
except ImportError:
    pd = None

# Ensure project root is on sys.path for Streamlit Cloud
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)
# Also add parent dir as fallback for environments like Streamlit Cloud
PARENT_DIR = os.path.dirname(ROOT_DIR)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)
from src.pptx2md.extract import extract_pptx_to_docs
from src.pptx2md.markdown import docs_to_markdown
from src.pptx2md.options import ExtractOptions
from src.pptx2md.translate import translate_markdown, TranslationConfig
from src.pptx2md.ppt_generator import create_translated_presentation_v2

# 용어집 파일 제한 설정
MAX_GLOSSARY_ENTRIES = 500  # 최대 용어 개수
MAX_FILE_SIZE_MB = 5        # 최대 파일 크기 (MB)
MAX_TERM_LENGTH = 100       # 개별 용어 최대 길이
TMP_DIR = os.path.join(ROOT_DIR, "tmp")
os.makedirs(TMP_DIR, exist_ok=True)

def _load_glossary_from_bytes(file_name: str, file_bytes: bytes) -> dict | None:
    """바이트 데이터를 기반으로 용어집을 파싱합니다."""
    if not file_bytes:
        st.warning("용어집 파일이 비어있습니다.")
        return None

    file_size_mb = len(file_bytes) / (1024 * 1024)
    if file_size_mb > MAX_FILE_SIZE_MB:
        st.error(f"파일 크기가 너무 큽니다. 최대 {MAX_FILE_SIZE_MB}MB까지 지원됩니다. (현재: {file_size_mb:.1f}MB)")
        return None

    file_extension = file_name.lower().split('.')[-1]

    if file_extension == 'json':
        try:
            text = file_bytes.decode('utf-8')
        except UnicodeDecodeError as e:
            st.error(f"JSON 파일을 UTF-8로 해석할 수 없습니다: {str(e)}")
            return None
        try:
            glossary = json.loads(text)
            if not isinstance(glossary, dict):
                st.error("JSON 파일은 딕셔너리 형태여야 합니다.")
                return None
            return _validate_glossary(glossary)
        except json.JSONDecodeError as e:
            st.error(f"JSON 파일 형식이 올바르지 않습니다: {str(e)}")
            return None

    elif file_extension in ['xlsx', 'xls']:
        if pd is None:
            st.error("엑셀 파일 처리를 위해 pandas가 필요합니다. requirements.txt를 확인해주세요.")
            return None
        try:
            df = pd.read_excel(io.BytesIO(file_bytes))
        except Exception as e:
            st.error(f"엑셀 파일 읽기 오류: {str(e)}")
            return None

        if len(df.columns) < 2:
            st.error("엑셀 파일은 최소 2개의 컬럼이 필요합니다 (원문, 번역)")
            return None

        if len(df) > MAX_GLOSSARY_ENTRIES:
            st.error(f"용어 개수가 너무 많습니다. 최대 {MAX_GLOSSARY_ENTRIES}개까지 지원됩니다. (현재: {len(df)}개)")
            return None

        glossary = {}
        skipped_rows = 0

        for idx, row in df.iterrows():
            source = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            target = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""

            if not source or not target:
                skipped_rows += 1
                continue

            if len(source) > MAX_TERM_LENGTH or len(target) > MAX_TERM_LENGTH:
                st.warning(f"행 {idx + 2}: 용어가 너무 깁니다 (최대 {MAX_TERM_LENGTH}자). 건너뜁니다.")
                skipped_rows += 1
                continue

            glossary[source] = target

        if skipped_rows > 0:
            st.info(f"{skipped_rows}개 행이 건너뛰어졌습니다 (빈 값 또는 너무 긴 용어)")

        return _validate_glossary(glossary)

    else:
        st.error("지원하지 않는 파일 형식입니다. JSON 또는 엑셀 파일을 업로드해주세요.")
        return None


def load_glossary_from_file(uploaded_file) -> dict:
    """업로드된 파일에서 용어집을 로드합니다. JSON과 엑셀 파일을 모두 지원합니다."""
    if uploaded_file is None:
        return None
    file_bytes = uploaded_file.getvalue()
    return _load_glossary_from_bytes(uploaded_file.name, file_bytes)

def _validate_glossary(glossary: dict) -> dict:
    """용어집 최종 검증"""
    if not glossary:
        st.warning("용어집이 비어있습니다.")
        return None
    
    if len(glossary) > MAX_GLOSSARY_ENTRIES:
        st.error(f"용어 개수가 너무 많습니다. 최대 {MAX_GLOSSARY_ENTRIES}개까지 지원됩니다.")
        return None
    
    # 중복 제거 및 통계
    original_count = len(glossary)
    glossary = {k.strip(): v.strip() for k, v in glossary.items() if k.strip() and v.strip()}
    
    if len(glossary) != original_count:
        st.info(f"중복 또는 빈 항목 제거됨: {original_count} → {len(glossary)}개")
    
    return glossary


def get_glossary_from_upload(uploaded_file):
    """업로드된 용어집을 캐싱하고 재사용합니다."""
    if uploaded_file is None:
        st.session_state.pop("cached_glossary", None)
        st.session_state.pop("cached_glossary_meta", None)
        return None

    file_bytes = uploaded_file.getvalue()
    file_hash = hashlib.md5(file_bytes).hexdigest()
    meta = {
        "name": uploaded_file.name,
        "hash": file_hash,
        "size": len(file_bytes),
    }

    if st.session_state.get("cached_glossary_meta") == meta:
        return st.session_state.get("cached_glossary")

    glossary = _load_glossary_from_bytes(uploaded_file.name, file_bytes)
    if glossary:
        st.session_state["cached_glossary"] = glossary
        st.session_state["cached_glossary_meta"] = meta
    else:
        st.session_state.pop("cached_glossary", None)
        st.session_state.pop("cached_glossary_meta", None)
    return glossary

load_dotenv()
st.set_page_config(page_title="PPT 번역 솔루션", layout="centered")

# 헤더 로고: 파일이 존재할 때만 표시해 레이아웃을 깨뜨리지 않는다
logo_path = os.path.join(ROOT_DIR, "assets", "ppt_logo.png")
header_cols = st.columns([1, 8]) if os.path.exists(logo_path) else None

if header_cols:
    with header_cols[0]:
        st.image(logo_path, width=80)
    with header_cols[1]:
        st.title("PPT 번역 솔루션")
else:
    st.title("PPT 번역 솔루션")

if "translation_logs" not in st.session_state:
    st.session_state.translation_logs = []
if "log_placeholder" not in st.session_state:
    st.session_state.log_placeholder = None


def _render_logs():
    placeholder = st.session_state.get("log_placeholder")
    if placeholder is None:
        return

    if st.session_state.translation_logs:
        content = "<br/>".join(html.escape(msg) for msg in st.session_state.translation_logs)
        html_block = (
            "<div style='max-height:320px; overflow-y:auto; background-color:#0f1116; "
            "padding:12px; border-radius:8px; font-family:monospace; font-size:13px;'>"
            f"{content}</div>"
        )
        placeholder.markdown(html_block, unsafe_allow_html=True)
    else:
        placeholder.markdown("_진행 로그가 여기에 표시됩니다._")


def append_log(message: str):
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.translation_logs.append(f"[{timestamp}] {message}")
    _render_logs()


def reset_logs():
    st.session_state.translation_logs = []
    _render_logs()

with st.sidebar:
    st.header("기능 선택")
    
    # 페이지 네비게이션 버튼
    current_page = st.session_state.get("current_page", "extract")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📄 텍스트 추출", use_container_width=True, type="primary" if current_page == "extract" else "secondary"):
            st.session_state.current_page = "extract"
            st.rerun()
    
    with col2:
        if st.button("🌐 PPT 번역", use_container_width=True, type="primary" if current_page == "translate" else "secondary"):
            st.session_state.current_page = "translate"
            st.rerun()
    
    st.divider()
    
    # 현재 페이지에 따른 옵션 표시
    if current_page == "extract":
        st.subheader("텍스트 추출 옵션")
        with_notes = st.checkbox("발표자 노트 포함", value=False)
        # UI에서는 한국어로 표시하되 실제 값은 영어로 매핑
        figures_display = st.selectbox("그림 처리", ["플레이스홀더", "생략"], index=0)
        figures_map = {"플레이스홀더": "placeholder", "생략": "omit"}
        figures = figures_map[figures_display]
        
        charts_display = st.selectbox("차트 처리", ["레이블", "플레이스홀더", "생략"], index=0)
        charts_map = {"레이블": "labels", "플레이스홀더": "placeholder", "생략": "omit"}
        charts = charts_map[charts_display]
        
        # 기본값 설정 (번역 페이지에서 사용할 때)
        model = "gpt-4o-mini"
        default_prompt = """당신은 시니어 번역사입니다. PPT 번역 시:
- 원문 의미 유지하되 간결하게 번역
- 번역문이 원문보다 20% 이상 길어지지 않게 제한
- 자연스럽고 비즈니스에 적합한 표현 사용"""
        extra_prompt = default_prompt
        glossary_file = None
        
    elif current_page == "translate":
        st.subheader("번역 옵션")
        model = st.selectbox("OpenAI 모델", ["gpt-5", "gpt-4.1", "gpt-4.1-mini", "gpt-4o-mini", "gpt-5-nano"], index=3)
        default_prompt = """당신은 시니어 번역사입니다. PPT 번역 시:
- 원문 의미 유지하되 간결하게 번역
- 번역문이 원문보다 20% 이상 길어지지 않게 제한
- 자연스럽고 비즈니스에 적합한 표현 사용"""
        
        extra_prompt = st.text_area("번역 프롬프트", value=default_prompt, height=150, placeholder="톤, 스타일, 용어 규칙 등...")
        
        # 용어집 파일 업로더 - JSON과 엑셀 모두 지원
        glossary_file = st.file_uploader(
            "용어집 파일", 
            type=["json", "xlsx", "xls"],
            help="JSON 파일 또는 엑셀 파일을 업로드하세요. 엑셀의 경우 첫 번째 컬럼은 원문, 두 번째 컬럼은 번역어로 구성해주세요.",
            key="glossary_upload"
        )
        
        # 용어집 파일 제한사항 안내 (compact)
        st.markdown(f"""
        **용어집 파일 제한사항:**  
        • 최대 파일 크기: {MAX_FILE_SIZE_MB}MB  
        • 최대 용어 개수: {MAX_GLOSSARY_ENTRIES}개  
        • 개별 용어 최대 길이: {MAX_TERM_LENGTH}자
        """, unsafe_allow_html=True)
        
        # 용어집 미리보기
        if glossary_file:
            glossary_preview = get_glossary_from_upload(glossary_file)
            if glossary_preview:
                st.success(f"✅ 용어집 로드 완료: {len(glossary_preview)}개 항목")
                
                with st.expander("용어집 미리보기", expanded=False):
                    preview_items = list(glossary_preview.items())[:10]  # 처음 10개만 표시
                    for source, target in preview_items:
                        st.write(f"• `{source}` → `{target}`")
                    if len(glossary_preview) > 10:
                        st.write(f"... 외 {len(glossary_preview) - 10}개 항목")
        else:
            get_glossary_from_upload(None)
            
        # 기본값 설정 (텍스트 추출 페이지에서 사용할 때)
        with_notes = False
        figures = "placeholder"
        charts = "labels"

for k in ["uploaded_path", "docs", "markdown", "translated_md", "show_translation_tab", "output_pptx_path", "output_pptx_name"]:
    if k not in st.session_state:
        st.session_state[k] = None

if "last_action" not in st.session_state:
    st.session_state.last_action = None

if "current_page" not in st.session_state:
    st.session_state.current_page = "extract"


def run_action(action_type: str):
    """공통 액션 실행 함수"""
    try:
        if action_type == "translate_markdown":
            reset_logs()
            append_log("Markdown 번역 준비 중...")
            glossary = get_glossary_from_upload(glossary_file) if glossary_file else st.session_state.get("cached_glossary")
            if glossary:
                append_log(f"용어집 적용: {len(glossary)}개 항목")
            cfg = TranslationConfig(target_lang="en", glossary=glossary, extra_instructions=extra_prompt, model=model)

            start = time.time()
            append_log(f"Markdown 번역 요청 전송 — 글자 수 {len(st.session_state.markdown or ''):,}자")
            with st.spinner("번역 중..."):
                st.session_state.translated_md = translate_markdown(st.session_state.markdown, cfg)
                st.session_state.show_translation_tab = True
            elapsed = int(time.time() - start)
            append_log(f"Markdown 번역 완료 ({elapsed//60}분 {elapsed%60}초)")
            st.info(f"번역 소요 시간: {elapsed//60}분 {elapsed%60}초")
            st.session_state.last_action = "translate_markdown"
            st.rerun()

        elif action_type == "translate_ppt":
            reset_logs()
            append_log("PPT 번역 준비 중...")
            glossary = get_glossary_from_upload(glossary_file) if glossary_file else st.session_state.get("cached_glossary")
            if glossary:
                append_log(f"용어집 적용: {len(glossary)}개 항목")
            cfg = TranslationConfig(target_lang="en", glossary=glossary, extra_instructions=extra_prompt, model=model)

            base_name = os.path.splitext(st.session_state.get("uploaded_original_name") or os.path.basename(st.session_state.uploaded_path))[0]
            output_pptx = os.path.abspath(f"{base_name}_translated.pptx")
            if st.session_state.output_pptx_path and os.path.exists(st.session_state.output_pptx_path):
                try:
                    os.remove(st.session_state.output_pptx_path)
                except OSError:
                    pass

            start = time.time()
            slide_count = len(st.session_state.docs) if st.session_state.docs else "?"
            append_log(f"PPT 번역 및 생성 시작 — 대상 슬라이드 {slide_count}")
            with st.spinner("PPT 번역 및 생성 중..."):
                create_translated_presentation_v2(
                    st.session_state.uploaded_path,
                    output_pptx,
                    cfg,
                    progress_callback=append_log,
                )
            elapsed = int(time.time() - start)
            st.success(f"PPT 생성 완료! 소요 시간: {elapsed//60}분 {elapsed%60}초")
            append_log(f"PPT 번역 완료 ({elapsed//60}분 {elapsed%60}초)")

            st.session_state.output_pptx_path = output_pptx
            st.session_state.output_pptx_name = f"{base_name}_translated.pptx"
            st.session_state.last_action = "translate_ppt"
            st.rerun()

    except Exception as e:
        append_log(f"오류: {str(e)}")
        st.error(f"실행 중 오류가 발생했습니다: {str(e)}")


# 현재 페이지에 따른 메인 컨텐츠 표시
current_page = st.session_state.get("current_page", "extract")

if current_page == "extract":
    st.header("📄 PPT 텍스트 추출")
    st.write("PPT 파일에서 텍스트를 추출하여 Markdown 형식으로 변환하고, 필요시 번역할 수 있습니다.")
    
    uploaded = st.file_uploader("PPTX 파일 업로드", type=["pptx"], key="extract_uploader") 
    if uploaded:
        file_bytes = uploaded.getvalue()
        file_hash = hashlib.md5(file_bytes).hexdigest()
        meta = {
            "name": uploaded.name,
            "hash": file_hash,
            "size": len(file_bytes),
        }

        if st.session_state.get("uploaded_file_meta") != meta:
            if st.session_state.uploaded_path and os.path.exists(st.session_state.uploaded_path):
                try:
                    os.remove(st.session_state.uploaded_path)
                except OSError:
                    pass

            tmp_filename = f"{uuid.uuid4().hex}_{uploaded.name}"
            tmp_path = os.path.join(TMP_DIR, tmp_filename)
            with open(tmp_path, "wb") as f:
                f.write(file_bytes)

            st.session_state.uploaded_path = tmp_path
            st.session_state.uploaded_original_name = uploaded.name
            st.session_state.docs = None
            st.session_state.markdown = None
            st.session_state.translated_md = None
            st.session_state.output_pptx_path = None
            st.session_state.output_pptx_name = None
            st.session_state.uploaded_file_meta = meta
            reset_logs()

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Markdown 변환", use_container_width=True, disabled=not st.session_state.uploaded_path):
            reset_logs()
            append_log("슬라이드에서 텍스트 추출 시작")
            opts = ExtractOptions(with_notes=with_notes, figures=figures, charts=charts)
            docs = extract_pptx_to_docs(st.session_state.uploaded_path, opts)
            st.session_state.docs = docs
            st.session_state.markdown = docs_to_markdown(docs, opts)
            append_log(f"Markdown 생성 완료 — 슬라이드 {len(docs)}개")

    with col2:
        if st.button("번역 (Markdown)", use_container_width=True, disabled=not st.session_state.markdown):
            run_action("translate_markdown")

elif current_page == "translate":
    st.header("🌐 번역된 PPT 생성")
    st.write("원본 PPT 파일의 디자인을 유지하면서 내부 텍스트만 번역된 새로운 PPT 파일을 생성합니다.")
    
    uploaded = st.file_uploader("PPTX 파일 업로드", type=["pptx"], key="translate_uploader") 
    if uploaded:
        file_bytes = uploaded.getvalue()
        file_hash = hashlib.md5(file_bytes).hexdigest()
        meta = {
            "name": uploaded.name,
            "hash": file_hash,
            "size": len(file_bytes),
        }

        if st.session_state.get("uploaded_file_meta") != meta:
            if st.session_state.uploaded_path and os.path.exists(st.session_state.uploaded_path):
                try:
                    os.remove(st.session_state.uploaded_path)
                except OSError:
                    pass

            tmp_filename = f"{uuid.uuid4().hex}_{uploaded.name}"
            tmp_path = os.path.join(TMP_DIR, tmp_filename)
            with open(tmp_path, "wb") as f:
                f.write(file_bytes)

            st.session_state.uploaded_path = tmp_path
            st.session_state.uploaded_original_name = uploaded.name
            st.session_state.docs = None
            st.session_state.markdown = None
            st.session_state.translated_md = None
            st.session_state.output_pptx_path = None
            st.session_state.output_pptx_name = None
            st.session_state.uploaded_file_meta = meta
            reset_logs()

    if st.button("번역된 PPT 생성", use_container_width=True, disabled=not st.session_state.uploaded_path):
        run_action("translate_ppt")

# 페이지별 결과 표시
if current_page == "extract":
    # 텍스트 추출 페이지의 미리보기
    if st.session_state.markdown:
        st.divider()
        # Auto-switch to translation tab if translation exists
        default_tab = 1 if st.session_state.translated_md else 0
        tab1, tab2 = st.tabs(["Markdown 미리보기", "번역본 미리보기"])
        
        with tab1:
            st.code(st.session_state.markdown, language="markdown", height=400)
            st.download_button("Markdown 다운로드", st.session_state.markdown.encode("utf-8"), 
                              os.path.splitext(os.path.basename(st.session_state.uploaded_path))[0] + ".md")
        
        with tab2:
            if st.session_state.translated_md:
                st.code(st.session_state.translated_md, language="markdown", height=400)
            else:
                st.info("번역을 먼저 실행해주세요.")

elif current_page == "translate":
    # PPT 번역 페이지의 다운로드 버튼
    if st.session_state.output_pptx_path and os.path.exists(st.session_state.output_pptx_path):
        st.divider()
        with open(st.session_state.output_pptx_path, "rb") as f:
            st.download_button(
                "번역된 PPT 다운로드",
                data=f.read(),
                file_name=st.session_state.output_pptx_name or os.path.basename(st.session_state.output_pptx_path),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

# 공통 진행 로그 (모든 페이지에서 표시)
st.divider()
with st.container():
    st.subheader("진행 로그")
    st.caption("로그가 길어지면 스크롤하여 확인하세요.")
    st.session_state.log_placeholder = st.empty()
    _render_logs()
