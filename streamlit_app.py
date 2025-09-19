from __future__ import annotations

import io
import os
import time

import streamlit as st
from dotenv import load_dotenv

from src.pptx2md.extract import extract_pptx_to_docs
from src.pptx2md.markdown import docs_to_markdown
from src.pptx2md.options import ExtractOptions

load_dotenv()

st.set_page_config(page_title="PPTX → Markdown", page_icon="📝", layout="centered")
st.title("PPTX → Markdown 변환")

# Sidebar controls
with st.sidebar:
    st.header("옵션")
    with_notes = st.checkbox("발표자 노트 포함", value=False)
    figures = st.selectbox("그림 처리", options=["placeholder", "omit"], index=0)
    charts = st.selectbox("차트 처리", options=["labels", "placeholder", "omit"], index=0)
    st.divider()
    st.subheader("번역(영어)")
    do_translate = st.checkbox("번역 및 PPT 재생성", value=False)
    glossary_file = st.file_uploader("용어집(JSON, 선택)", type=["json"], key="glossary")
    extra_prompt = st.text_area("번역 참고 프롬프트(선택)", height=120, placeholder="Tone, style, terminology rules...")

# Session state for uploaded file path and outputs
if "uploaded_path" not in st.session_state:
    st.session_state.uploaded_path = None
if "docs" not in st.session_state:
    st.session_state.docs = None
if "markdown" not in st.session_state:
    st.session_state.markdown = None

# Upload control (does not auto-process)
uploaded = st.file_uploader("PPTX 파일 업로드", type=["pptx"], key="pptx") 
if uploaded is not None:
    tmp_path = os.path.abspath("_tmp_upload.pptx")
    with open(tmp_path, "wb") as f:
        f.write(uploaded.read())
    st.session_state.uploaded_path = tmp_path
    st.success("업로드 완료. 아래 버튼으로 변환을 시작하세요.")

# Action buttons
col1, col2 = st.columns(2)
with col1:
    if st.button("Markdown 변환", use_container_width=True, disabled=st.session_state.uploaded_path is None):
        options = ExtractOptions(with_notes=with_notes, figures=figures, charts=charts)  # type: ignore
        docs = extract_pptx_to_docs(st.session_state.uploaded_path, options)
        md_text = docs_to_markdown(docs, options)
        st.session_state.docs = docs
        st.session_state.markdown = md_text
        st.success("변환 완료")
with col2:
    if st.button("번역 + PPT 생성", use_container_width=True, disabled=not do_translate or st.session_state.markdown is None):
        from src.pptx2md.translate import translate_markdown, TranslationConfig
        from src.pptx2md.reinsert import create_translated_copy, OverflowPolicy
        import json

        glossary = None
        if glossary_file is not None:
            try:
                glossary = json.loads(glossary_file.getvalue().decode("utf-8"))
            except Exception:
                st.error("용어집 JSON 파싱 실패")
        cfg = TranslationConfig(target_lang="en", glossary=glossary, extra_instructions=extra_prompt)
        # timer UI
        start = time.time()
        placeholder = st.empty()
        with st.spinner("번역 중..."):
            translated_md = translate_markdown(st.session_state.markdown, cfg)
            while False:
                # kept for future streaming; currently spinner handles busy state
                pass
        elapsed = int(time.time() - start)
        mm, ss = divmod(elapsed, 60)
        placeholder.info(f"번역 소요 시간: {mm}분 {ss}초")

        st.download_button(
            "번역된 Markdown 다운로드",
            data=translated_md.encode("utf-8"),
            file_name=(os.path.splitext(os.path.basename(st.session_state.uploaded_path))[0] + ".en.md"),
            mime="text/markdown",
        )
        # 재삽입(현재는 간이 매핑; 모듈 경계는 reinsert.py에 분리)
        from src.pptx2md.models import SlideDoc, TextBlock
        translated_docs = []
        for d in st.session_state.docs:
            new_blocks = []
            for b in d.blocks:
                if isinstance(b, TextBlock):
                    new_blocks.append(TextBlock(shape_id=b.shape_id, lines=b.lines, indent_levels=b.indent_levels))
                else:
                    new_blocks.append(b)
            translated_docs.append(SlideDoc(slide_index=d.slide_index, title=d.title, blocks=new_blocks))
        out_pptx = os.path.abspath("translated.pptx")
        policy = OverflowPolicy()
        create_translated_copy(st.session_state.uploaded_path, translated_docs, out_pptx, policy)
        with open(out_pptx, "rb") as f:
            st.download_button("번역된 PPTX 다운로드", data=f.read(), file_name=(os.path.splitext(os.path.basename(st.session_state.uploaded_path))[0] + ".en.pptx"), mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# Output preview
if st.session_state.markdown:
    st.subheader("Markdown 미리보기")
    st.code(st.session_state.markdown, language="markdown")
