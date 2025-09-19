from __future__ import annotations
import os, time, json
import streamlit as st
from dotenv import load_dotenv
from src.pptx2md.extract import extract_pptx_to_docs
from src.pptx2md.markdown import docs_to_markdown
from src.pptx2md.options import ExtractOptions
from src.pptx2md.translate import translate_markdown, TranslationConfig

load_dotenv()
st.set_page_config(page_title="PPTX → Markdown", layout="centered")
st.title("PPTX → Markdown 변환")

with st.sidebar:
    st.header("옵션")
    with_notes = st.checkbox("발표자 노트 포함", value=False)
    figures = st.selectbox("그림 처리", ["placeholder", "omit"], index=0)
    charts = st.selectbox("차트 처리", ["labels", "placeholder", "omit"], index=0)
    st.divider()
    st.subheader("번역(영어)")
    model = st.selectbox("OpenAI 모델", ["gpt-5", "gpt-4.1", "gpt-4.1-mini", "gpt-4o-mini", "gpt-5-nano"], index=3)
    extra_prompt = st.text_area("번역 참고 프롬프트", height=120, placeholder="Tone, style, terminology rules...")
    glossary_file = st.file_uploader("용어집(JSON)", type=["json"])  # 선택

for k in ["uploaded_path", "docs", "markdown", "translated_md", "show_translation_tab"]:
    if k not in st.session_state:
        st.session_state[k] = None

uploaded = st.file_uploader("PPTX 파일 업로드", type=["pptx"]) 
if uploaded:
    tmp_path = os.path.abspath(f"_tmp_{uploaded.name}")
    with open(tmp_path, "wb") as f: f.write(uploaded.read())
    st.session_state.uploaded_path = tmp_path

col1, col2 = st.columns(2)
with col1:
    if st.button("Markdown 변환", use_container_width=True, disabled=not st.session_state.uploaded_path):
        opts = ExtractOptions(with_notes=with_notes, figures=figures, charts=charts)
        docs = extract_pptx_to_docs(st.session_state.uploaded_path, opts)
        st.session_state.docs = docs
        st.session_state.markdown = docs_to_markdown(docs, opts)

with col2:
    if st.button("번역 (Markdown 기반)", use_container_width=True, disabled=not st.session_state.markdown):
        glossary = json.loads(glossary_file.getvalue()) if glossary_file else None
        cfg = TranslationConfig(target_lang="en", glossary=glossary, extra_instructions=extra_prompt, model=model)
        start = time.time()
        with st.spinner("번역 중..."):
            st.session_state.translated_md = translate_markdown(st.session_state.markdown, cfg)
            st.session_state.show_translation_tab = True
        elapsed = int(time.time() - start)
        st.info(f"번역 소요 시간: {elapsed//60}분 {elapsed%60}초")
        st.rerun()  # Force refresh to switch to translation tab

# Tabbed preview sections
if st.session_state.markdown:
    # Auto-switch to translation tab if translation exists
    default_tab = 1 if st.session_state.translated_md else 0
    tab1, tab2 = st.tabs(["Markdown 미리보기", "번역본 미리보기"])
    
    with tab1:
        st.code(st.session_state.markdown, language="markdown")
        st.download_button("Markdown 다운로드", st.session_state.markdown.encode("utf-8"), 
                          os.path.splitext(os.path.basename(st.session_state.uploaded_path))[0] + ".md")
    
    with tab2:
        if st.session_state.translated_md:
            st.code(st.session_state.translated_md, language="markdown")
        else:
            st.info("번역을 먼저 실행해주세요.")
