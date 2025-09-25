from __future__ import annotations
import os, time, json, io, uuid, hashlib
from contextlib import nullcontext
from pptx import Presentation
import streamlit as st
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
from src.pptx2md.translate import TranslationConfig
from src.pptx2md.ppt_generator import create_translated_presentation_v2, compress_images_in_presentation, optimize_pptx_media_zip

# 용어집 파일 제한 설정
MAX_GLOSSARY_ENTRIES = 500  # 최대 용어 개수
MAX_FILE_SIZE_MB = 5        # 최대 파일 크기 (MB)
MAX_TERM_LENGTH = 100       # 개별 용어 최대 길이
TMP_DIR = os.path.join(ROOT_DIR, "tmp")
os.makedirs(TMP_DIR, exist_ok=True)
OUTPUT_DIR = os.path.join(ROOT_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

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

st.set_page_config(page_title="PPT 번역캣", layout="centered")

# Progress bar를 전체 영역 너비로 확장
st.markdown(
    """
    <style>
    /* Progress bar 컨테이너를 전체 너비로 설정 */
    .stProgress {
        width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    
    /* Progress bar 자체를 전체 너비로 설정 */
    div[data-testid="stProgress"] {
        width: 100% !important;
        padding: 0 !important;
        margin-top: 0.25rem;
    }
    
    /* Progress bar 내부 요소들 전체 너비로 설정 */
    div[data-testid="stProgress"] > div {
        width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    
    /* Progress bar의 실제 바 요소 */
    div[data-testid="stProgress"] > div > div {
        width: 100% !important;
    }
    
    /* 모든 progress 관련 요소 */
    [data-testid="stProgress"] * {
        width: 100% !important;
        max-width: 100% !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 헤더 로고: 파일이 존재할 때만 표시해 레이아웃을 깨뜨리지 않는다
logo_path = os.path.join(ROOT_DIR, "assets", "ppt_logo.png")
header_cols = st.columns([1, 8]) if os.path.exists(logo_path) else None

if header_cols:
    with header_cols[0]:
        st.image(logo_path, width=80)
    with header_cols[1]:
        st.title("PPT 번역캣")
else:
    st.title("PPT 번역캣")

if "last_status" not in st.session_state:
    st.session_state.last_status = None


def _set_status(kind: str | None, message: str | None = None) -> None:
    if kind is None:
        st.session_state.last_status = None
        return
    st.session_state.last_status = {"type": kind, "message": message or ""}

with st.sidebar:
    st.header("기능 선택")
    
    # 페이지 네비게이션 버튼
    current_page = st.session_state.get("current_page", "extract")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("📄 텍스트 추출", use_container_width=True, type="primary" if current_page == "extract" else "secondary"):
            st.session_state.current_page = "extract"
            st.rerun()
    
    with col2:
        if st.button("🌐 PPT 번역", use_container_width=True, type="primary" if current_page == "translate" else "secondary"):
            st.session_state.current_page = "translate"
            st.rerun()

    with col3:
        if st.button("🖼️ 이미지 최적화", use_container_width=True, type="primary" if current_page == "optimize_images" else "secondary"):
            st.session_state.current_page = "optimize_images"
            st.rerun()
    
    st.divider()
    
    # 현재 페이지에 따른 옵션 표시
    if current_page == "extract":
        st.subheader("텍스트 추출 옵션")
        with_notes = st.checkbox("발표자 노트 포함", value=False)
        # UI에서는 한국어로 표시하되 실제 값은 영어로 매핑
        figures_display = st.selectbox("그림 처리", ["플레이스홀더", "생략"], index=1)
        figures_map = {"플레이스홀더": "placeholder", "생략": "omit"}
        figures = figures_map[figures_display]
        
        charts_display = st.selectbox("차트 처리", ["제목만", "플레이스홀더", "생략"], index=0)
        charts_map = {"제목만": "labels", "플레이스홀더": "placeholder", "생략": "omit"}
        charts = charts_map[charts_display]
        
        # (텍스트 추출 페이지에서는 번역 관련 입력을 사용하지 않음)
        
    elif current_page == "translate":
        st.subheader("번역 옵션")
        language_choices = [
            ("자동 감지", "auto"),
            ("한국어", "ko"),
            ("영어", "en"),
            ("일본어", "ja"),
            ("중국어", "zh"),
        ]
        language_map = {label: code for label, code in language_choices}
        source_labels = [label for label, _ in language_choices]
        target_labels = [label for label, code in language_choices if code != "auto"]

        default_source_code = st.session_state.get("source_lang", "auto")
        default_target_code = st.session_state.get("target_lang", "en")

        source_index = next((idx for idx, (_, code) in enumerate(language_choices) if code == default_source_code), 0)
        target_index = next((idx for idx, label in enumerate(target_labels) if language_map[label] == default_target_code), 1 if "영어" in target_labels else 0)

        source_lang_label = st.selectbox("소스 언어", source_labels, index=source_index)
        target_lang_label = st.selectbox("타겟 언어", target_labels, index=target_index)

        source_lang = language_map[source_lang_label]
        target_lang = language_map[target_lang_label]
        language_pair_display = f"{source_lang_label}→{target_lang_label}"

        st.session_state.source_lang = source_lang
        st.session_state.target_lang = target_lang
        st.session_state.language_pair_display = language_pair_display

        if source_lang == "auto":
            rule_one = f"문장별 언어를 판별하여 {target_lang_label}로 번역."
        elif source_lang == target_lang:
            rule_one = f"{target_lang_label} 문장은 의미를 유지하며 자연스럽게 다듬기."
        else:
            rule_one = f"{source_lang_label} 원문을 {target_lang_label}로 번역."

        model_options = ["gpt-4o-mini", "gpt-4.1", "gpt-5"]
        default_model = st.session_state.get("selected_model", model_options[2])
        if default_model not in model_options:
            default_model = model_options[0]
        model_index = model_options.index(default_model)
        model = st.selectbox("OpenAI 모델", model_options, index=model_index)
        st.session_state.selected_model = model
        default_prompt = f"""1. {rule_one}
2. 이미 {target_lang_label}로 된 문장/용어는 그대로 두거나 자연스럽게 다듬기만.
3. 의미 유지 + 간결·명료, 길이는 원문 120 % 이내.
4. 용어집 우선, 고유명사는 원형 유지.
5. 개발·마케팅 실무자가 읽기 쉬운 자연스러운 표현 사용."""
        
        extra_prompt = st.text_area("번역 프롬프트", value=default_prompt, height=150, placeholder="톤, 스타일, 용어 규칙 등...")
        
        # 이미지 최적화 옵션
        st.markdown("**이미지 품질 최적화**")
        enable_img_opt = st.checkbox("이미지 용량 최적화 활성화", value=False, help="고해상도 이미지를 다운스케일/재압축하여 PPT 용량을 줄입니다.")
        col_q, col_px = st.columns(2)
        with col_q:
            img_quality = st.slider("JPEG 품질(%)", min_value=10, max_value=95, value=70, step=5, help="낮을수록 용량이 줄지만 품질이 떨어집니다.")
        with col_px:
            img_max_px = st.slider("최대 긴 변(px)", min_value=800, max_value=4096, value=1920, step=160, help="이 값보다 큰 이미지를 비율 유지하며 축소합니다.")

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

    elif current_page == "optimize_images":
        st.subheader("이미지 최적화 옵션")
        img_quality = st.slider("JPEG 품질(%)", min_value=10, max_value=95, value=70, step=5)
        img_max_px = st.slider("최대 긴 변(px)", min_value=800, max_value=4096, value=1920, step=160)

for k in ["uploaded_path", "docs", "markdown", "output_pptx_path", "output_pptx_name"]:
    if k not in st.session_state:
        st.session_state[k] = None

if "last_action" not in st.session_state:
    st.session_state.last_action = None

if "current_page" not in st.session_state:
    st.session_state.current_page = "extract"


def run_action(action_type: str, *, progress_slot=None):
    """공통 액션 실행 함수"""
    progress_label = progress_bar = None
    try:
        _set_status(None)
        if progress_slot is not None:
            progress_container = progress_slot.container()
            progress_label = progress_container.empty()
            progress_bar = progress_container.progress(0)

        def _set_progress(percent: float, message: str | None = None) -> None:
            if progress_bar is None:
                return
            pct = int(max(0.0, min(100.0, percent)))
            progress_bar.progress(pct)
            if progress_label is not None:
                if message:
                    progress_label.markdown(f"**{message}**")
                else:
                    progress_label.empty()

        _set_progress(0, "준비 중...")

        if action_type == "translate_ppt":
            glossary = get_glossary_from_upload(glossary_file) if glossary_file else st.session_state.get("cached_glossary")
            # 이미지 최적화 설정 가져오기 (전역 변수가 아니므로 함수 내에서 재참조)
            enable_img_opt = globals().get('enable_img_opt', False)
            img_quality = globals().get('img_quality', 70)
            img_max_px = globals().get('img_max_px', 1920)
            _set_progress(10, "용어집 적용 중...")

            cfg = TranslationConfig(
                source_lang=st.session_state.get("source_lang", "auto"),
                target_lang=st.session_state.get("target_lang", "en"),
                glossary=glossary,
                extra_instructions=extra_prompt,
                model=model,
            )

            original_name = st.session_state.get("uploaded_original_name") or os.path.basename(st.session_state.uploaded_path)
            base_name, base_ext = os.path.splitext(original_name)
            file_ext = base_ext if base_ext else ".pptx"
            lang_code = st.session_state.get("target_lang", "en")
            lang_tag = {"ko": "KR", "en": "EN", "ja": "JA", "zh": "ZH"}.get(lang_code, str(lang_code).upper())
            file_name = f"Trans_{lang_tag}_{base_name}{file_ext}"
            output_pptx = os.path.abspath(os.path.join(OUTPUT_DIR, file_name))
            if st.session_state.output_pptx_path and os.path.exists(st.session_state.output_pptx_path):
                try:
                    os.remove(st.session_state.output_pptx_path)
                except OSError:
                    pass

            start = time.time()

            def _on_progress(payload: dict[str, float | str]) -> None:
                ratio = float(payload.get("ratio", 0.0)) if payload else 0.0
                message = str(payload.get("message", "진행 중..."))
                _set_progress(ratio * 100, message)

            spinner_cm = st.spinner("OpenAI 번역 처리 중...") if progress_slot is not None else nullcontext()
            with spinner_cm:
                # 1) (선택) 번역 시작 전 입력 PPT를 이미지 최적화
                input_path_for_translation = st.session_state.uploaded_path
                img_optimization_stats = None
                if enable_img_opt:
                    _set_progress(12, "이미지 최적화 준비...")
                    preopt_path = os.path.join(TMP_DIR, f".__preopt_{uuid.uuid4().hex}.pptx")
                    def _on_preopt(payload: dict[str, float | str]) -> None:
                        ratio = float(payload.get("ratio", 0.0)) if payload else 0.0
                        message = f"이미지 최적화 — {payload.get('message','')}"
                        # 12%~28% 구간에서 최적화 진행률 표시
                        _set_progress(12 + int(min(16.0, ratio * 16.0)), message)
                    try:
                        img_optimization_stats = optimize_pptx_media_zip(st.session_state.uploaded_path, preopt_path, quality=img_quality, max_px=img_max_px, progress_cb=_on_preopt)
                        # 결과 검증
                        try:
                            Presentation(preopt_path)
                            input_path_for_translation = preopt_path
                        except Exception:
                            # 검증 실패 시 원본으로 진행
                            try:
                                os.remove(preopt_path)
                            except Exception:
                                pass
                    except PermissionError:
                        _set_progress(14, "이미지 최적화 실패(잠금/권한). 원본으로 진행합니다.")
                    except Exception:
                        _set_progress(14, "이미지 최적화 실패. 원본으로 진행합니다.")

                # 2) 번역 실행 (28% 이후는 기존 비율 콜백 사용)
                stats = create_translated_presentation_v2(
                    input_path_for_translation,
                    output_pptx,
                    cfg,
                    progress_callback=_on_progress,
                )

            elapsed = int(time.time() - start)
            _set_progress(100, "PPT 번역 완료")

            st.session_state.output_pptx_path = output_pptx
            st.session_state.output_pptx_name = file_name
            st.session_state.last_action = "translate_ppt"
            # 상세 로그 메시지 (구조화된 형태)
            try:
                output_size_mb = os.path.getsize(output_pptx) / (1024 * 1024)
                input_size_mb = os.path.getsize(st.session_state.uploaded_path) / (1024 * 1024) if st.session_state.uploaded_path else 0
                glossary_count = len(glossary) if isinstance(glossary, dict) else 0
                glossary_part = f"{glossary_count}항목 적용" if glossary_count > 0 else "없음"
                model_name = model if 'model' in locals() else getattr(cfg, 'model', 'unknown')
                stats = stats or {}
                slide_count = stats.get("slides") if isinstance(stats, dict) else 0
                word_count = stats.get("word_count") if isinstance(stats, dict) else 0
                language_pair = st.session_state.get("language_pair_display", "")
                
                # 이미지 최적화 정보 추가
                img_info = ""
                if enable_img_opt and img_optimization_stats:
                    img_optimized = img_optimization_stats.get("optimized", 0)
                    img_candidates = img_optimization_stats.get("media", 0)
                    img_saved_mb = img_optimization_stats.get("bytes_saved", 0) / (1024 * 1024)
                    img_info = f"\n\n🖼️ 이미지 최적화\n• 후보: {img_candidates}개 → 성공: {img_optimized}개\n• 용량 절감: {img_saved_mb:.1f}MB"
                
                reduction_pct = ((input_size_mb - output_size_mb) / input_size_mb * 100) if input_size_mb > 0 else 0
                
                msg = f"""✅ PPT 번역 완료

📊 번역 정보
• 언어: {language_pair}
• 모델: {model_name}
• 용어집: {glossary_part}
• 슬라이드: {slide_count}개
• 번역 단어: {word_count:,}개{img_info}

📁 파일 정보
• 입력: {input_size_mb:.1f}MB
• 출력: {output_size_mb:.1f}MB ({reduction_pct:.0f}% 절감)
• 파일명: {st.session_state.output_pptx_name}

⏱️ 소요 시간: {elapsed//60}분 {elapsed%60}초"""
            except Exception:
                msg = f"✅ PPT 번역 완료\n\n⏱️ 소요 시간: {elapsed//60}분 {elapsed%60}초"
            _set_status("success", msg)
            st.rerun()

    except Exception as e:
        if progress_slot:
            progress_slot.empty()
        _set_status("error", f"실패: {str(e)}")


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
            st.session_state.output_pptx_path = None
            st.session_state.output_pptx_name = None
            st.session_state.uploaded_file_meta = meta

    extract_clicked = st.button("Markdown 변환", use_container_width=True, disabled=not st.session_state.uploaded_path)
    extract_progress_slot = st.empty()
    if extract_clicked:
        _set_status(None)
        progress_container = extract_progress_slot.container()
        status_placeholder = progress_container.empty()
        progress_bar = progress_container.progress(0)

        status_placeholder.markdown("**슬라이드 분석 중...**")
        progress_bar.progress(10)
        opts = ExtractOptions(with_notes=with_notes, figures=figures, charts=charts)
        docs = extract_pptx_to_docs(st.session_state.uploaded_path, opts)

        status_placeholder.markdown("**Markdown 생성 중...**")
        progress_bar.progress(60)
        st.session_state.docs = docs
        st.session_state.markdown = docs_to_markdown(docs, opts)

        status_placeholder.markdown("**Markdown 변환 완료**")
        progress_bar.progress(100)
        # 상세 로그 메시지
        try:
            md_text = st.session_state.markdown or ""
            md_len = len(md_text)
            md_lines = (md_text.count("\n") + 1) if md_text else 0
            meta = st.session_state.get("uploaded_file_meta", {}) or {}
            src_name = meta.get("name") or (os.path.basename(st.session_state.uploaded_path) if st.session_state.uploaded_path else "")
            src_size_mb = (meta.get("size", 0) / (1024 * 1024)) if meta else (os.path.getsize(st.session_state.uploaded_path) / (1024 * 1024) if st.session_state.uploaded_path else 0)
            notes_label = "포함" if with_notes else "미포함"
            msg = (
                f"Markdown 변환 완료 — 슬라이드 {len(docs)}개, "
                f"옵션: 노트 {notes_label} / 그림 {figures_display} / 차트 {charts_display}, "
                f"MD {md_len:,}자·{md_lines:,}라인, 원본 '{src_name}'({src_size_mb:.1f}MB)"
            )
        except Exception:
            msg = f"Markdown 변환 완료 (슬라이드 {len(docs)}개)"
        _set_status("success", msg)

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
            st.session_state.output_pptx_path = None
            st.session_state.output_pptx_name = None
            st.session_state.uploaded_file_meta = meta

    generate_clicked = st.button("번역된 PPT 생성", use_container_width=True, disabled=not st.session_state.uploaded_path)
    ppt_progress_slot = st.empty()
    if generate_clicked:
        run_action("translate_ppt", progress_slot=ppt_progress_slot)

elif current_page == "optimize_images":
    st.header("🖼️ PPT 이미지 최적화")
    st.write("PPT 내부 이미지를 다운스케일/재압축하여 용량을 줄입니다. 텍스트는 변경하지 않습니다.")

    uploaded = st.file_uploader("PPTX 파일 업로드", type=["pptx"], key="opt_uploader")
    opt_progress = st.empty()
    optimize_clicked = st.button("이미지 최적화", use_container_width=True, disabled=not uploaded)
    if uploaded and optimize_clicked:
        tmp_filename = f"{uuid.uuid4().hex}_{uploaded.name}"
        tmp_path = os.path.join(TMP_DIR, tmp_filename)
        with open(tmp_path, "wb") as f:
            f.write(uploaded.getvalue())

        with st.spinner("이미지 최적화 중..."):
            progress_bar = opt_progress.progress(0)
            def _cb(payload: dict):
                pct = int(min(100, max(0, float(payload.get("ratio", 0.0)) * 100)))
                msg = str(payload.get("message", ""))
                progress_bar.progress(pct, text=msg)
            out_name = os.path.splitext(uploaded.name)[0] + "_optimized.pptx"
            out_path = os.path.join(OUTPUT_DIR, out_name)
            try:
                stats = optimize_pptx_media_zip(tmp_path, out_path, quality=img_quality, max_px=img_max_px, progress_cb=_cb)
                # 결과 검증
                verified = True
                try:
                    Presentation(out_path)
                except Exception:
                    verified = False
                if not verified:
                    st.error("최적화 결과 검증 실패. 결과 파일을 폐기하고 원본을 유지합니다.")
                    try:
                        os.remove(out_path)
                    except Exception:
                        pass
                else:
                    # 구조화된 성공 메시지
                    input_size_mb = os.path.getsize(tmp_path) / (1024 * 1024)
                    output_size_mb = os.path.getsize(out_path) / (1024 * 1024)
                    reduction_pct = ((input_size_mb - output_size_mb) / input_size_mb * 100) if input_size_mb > 0 else 0
                    img_optimized = stats.get("optimized", 0)
                    img_candidates = stats.get("media", 0)
                    img_saved_mb = stats.get("bytes_saved", 0) / (1024 * 1024)
                    
                    success_msg = f"""✅ 이미지 최적화 완료

🖼️ 이미지 최적화
• 후보: {img_candidates}개 → 성공: {img_optimized}개
• 용량 절감: {img_saved_mb:.1f}MB

📁 파일 정보
• 입력: {input_size_mb:.1f}MB
• 출력: {output_size_mb:.1f}MB ({reduction_pct:.0f}% 절감)
• 파일명: {out_name}"""
                    
                    st.success(success_msg)
                    with open(out_path, "rb") as f:
                        st.download_button("최적화 PPT 다운로드", data=f.read(), file_name=out_name, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            except PermissionError:
                st.error("파일 잠금/권한 문제로 저장에 실패했습니다. 열려있는 PPT를 닫고 다시 시도하세요.")
            except Exception as e:
                st.error(f"이미지 최적화 실패: {e}")

# 페이지별 결과 표시
if current_page == "extract":
    # 텍스트 추출 페이지의 미리보기
    if st.session_state.markdown:
        st.divider()
        st.code(st.session_state.markdown, language="markdown", height=400)
        st.download_button(
            "Markdown 다운로드",
            st.session_state.markdown.encode("utf-8"),
            os.path.splitext(os.path.basename(st.session_state.uploaded_path))[0] + ".md",
        )

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

status = st.session_state.get("last_status")
if status:
    st.divider()
    if status["type"] == "success":
        st.success(status["message"])
    elif status["type"] == "error":
        st.error(status["message"])
    else:
        st.info(status["message"])
