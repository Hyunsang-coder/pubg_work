# PPTX → Markdown Converter

Streamlit UI to extract text from PPTX and produce clean Markdown per slide.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env  # set APP_SECRET_KEY (optional)
```

## Run (Streamlit)

```bash
streamlit run streamlit_app.py
```

Then open the shown URL (usually http://localhost:8501), upload a .pptx, and download the Markdown.

## Options (UI)
- Include presenter notes
- Figures: placeholder or omit
- Charts: labels, placeholder, or omit

## Project Structure
- `src/pptx2md/` core extractor and markdown formatter
- `streamlit_app.py` Streamlit UI
- `app.py` Flask UI (optional, legacy)

## Notes
- Extraction to Markdown only. No translation or re-insertion.
- Local processing; files are not sent externally.
