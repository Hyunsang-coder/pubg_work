from __future__ import annotations

import io
import os

import streamlit as st
from dotenv import load_dotenv

from src.pptx2md.extract import extract_pptx_to_docs
from src.pptx2md.markdown import docs_to_markdown
from src.pptx2md.options import ExtractOptions

load_dotenv()

st.set_page_config(page_title="PPTX → Markdown", page_icon="📝", layout="centered")
st.title("PPTX → Markdown 변환")

with st.sidebar:
    st.header("옵션")
    with_notes = st.checkbox("발표자 노트 포함", value=False)
    figures = st.selectbox("그림 처리", options=["placeholder", "omit"], index=0)
    charts = st.selectbox("차트 처리", options=["labels", "placeholder", "omit"], index=0)

uploaded = st.file_uploader("PPTX 파일 업로드", type=["pptx"]) 

if uploaded is not None:
    # Save to temp in-memory buffer then to disk for python-pptx
    tmp_bytes = uploaded.read()
    tmp_path = os.path.abspath("_tmp_upload.pptx")
    with open(tmp_path, "wb") as f:
        f.write(tmp_bytes)

    options = ExtractOptions(with_notes=with_notes, figures=figures, charts=charts)  # type: ignore
    docs = extract_pptx_to_docs(tmp_path, options)
    md_text = docs_to_markdown(docs, options)

    st.success("변환 완료")
    st.download_button(
        "Markdown 다운로드",
        data=md_text.encode("utf-8"),
        file_name=os.path.splitext(uploaded.name)[0] + ".md",
        mime="text/markdown",
    )

    st.subheader("미리보기")
    st.code(md_text, language="markdown")
