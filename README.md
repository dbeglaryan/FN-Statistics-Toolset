# FrostNode | Statistics Toolset

Professional toolset that stays simple. Two modes (Simple / Advanced), charts you can download as PNG,
Excel-friendly outputs, and one-click exports.

## Highlights (new in v2)
- Download-All ZIP (charts + tables + summary.txt)
- Definitions & Tips panel
- Auto-narratives (histogram shape, regression conclusion, CI meaning)
- Import helpers (header, delimiter, coerce-to-numeric, drop log)
- Assumption checks: Q-Q plot, Shapiro–Wilk; Breusch–Pagan hint
- Transform toggles: log(y), log(x), both; back-prediction
- Spearman correlation option + 95% CI for r (Fisher z)
- Excel extras: absolute ranges toggle; Formula Sheet download (.txt)
- Export to **Excel**, **Word**, **PowerPoint** with one click
- Save/Load session (.json)

## Run locally
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app_streamlit.py
```

## Student workflow (super simple)
1. Upload CSV/Excel or choose a sample.  
2. Pick X (and Y for regression).  
3. In **Simple** mode: click Histogram or Scatter+Line → **Download PNG**.  
4. Need everything? Go to **Exports** → Download-All ZIP / Excel / Word / PowerPoint.
