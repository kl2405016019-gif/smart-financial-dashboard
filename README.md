# Smart Financial Dashboard

A lightweight Streamlit web app that turns the AC4313 Excel template into an interactive dashboard suitable for undergraduate learning or SME self-checks.

## Features
- Upload the provided Excel template (or start from blank).
- View key KPIs, ratio tables (5-year), and interactive trend charts.
- Download the computed ratios as CSV and an Excel report.
- Executive-summary helper prompts (auto-generated) for teaching/report writing.

## Run locally
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Notes
- Calculations are replicated in Python (the app does not rely on Excel formula recalculation).
- Years are treated as Y1..Y5 (you can rename them in the app).
