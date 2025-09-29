from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional

from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, send_file, flash

from src.pptx2md.extract import extract_pptx_to_docs
from src.pptx2md.markdown import docs_to_markdown
from src.pptx2md.options import ExtractOptions

load_dotenv()

UPLOAD_DIR = os.path.abspath("uploads")
OUTPUT_DIR = os.path.abspath("output")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.getenv("APP_SECRET_KEY", "dev-secret")


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("pptx_file")
        with_notes = request.form.get("with_notes") == "on"
        figures = request.form.get("figures", "placeholder")
        charts = request.form.get("charts", "labels")
        try:
            if not file or not file.filename.endswith(".pptx"):
                flash(".pptx 파일을 업로드하세요.")
                return redirect(url_for("index"))
            in_path = os.path.join(UPLOAD_DIR, file.filename)
            file.save(in_path)

            options = ExtractOptions(
                with_notes=with_notes,
                figures=figures,  # type: ignore
                charts=charts,    # type: ignore
            )
            docs = extract_pptx_to_docs(in_path, options)
            md = docs_to_markdown(docs, options)

            out_name = os.path.splitext(os.path.basename(in_path))[0] + ".md"
            out_path = os.path.join(OUTPUT_DIR, out_name)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(md)

            return redirect(url_for("result", filename=out_name))
        finally:
            pass
    return render_template("index.html")


@app.route("/result")
def result():
    filename = request.args.get("filename")
    if not filename:
        return redirect(url_for("index"))
    return render_template("result.html", filename=filename)


@app.route("/download/<path:filename>")
def download(filename: str):
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(path):
        flash("파일을 찾을 수 없습니다.")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    # Streamlit 등 다른 실행 환경에서 import될 때 signal 관련 오류가 발생하지 않도록
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", 5000)),
        debug=True,
        use_reloader=False,
    )
