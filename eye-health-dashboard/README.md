# Eye-Health Analytics Dashboard

Interactive dashboard exploring lifestyle, screen habits, and vision across
10,000 respondents. This folder ships **two ways to view the same data**:

| Version | Entry point | Stack |
|---------|-------------|-------|
| **Streamlit app** (deployable) | `app.py` | Python · Streamlit · Plotly |
| **Static HTML dashboard** | `index.html` | Plotly.js · noUiSlider |

Both read the same auto-generated `data.js` (10,000 rows × 11 attributes), so
there is a single source of truth.

## Run the Streamlit app locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Opens at http://localhost:8501.

## Deploy to Streamlit Community Cloud

1. Push this project to a public GitHub repo. Make sure `app.py`, `data.js`,
   `requirements.txt`, and `.streamlit/config.toml` are committed
   (`node_modules/` is git-ignored — it is not needed for deploy).
2. Go to https://share.streamlit.io → **Create app** → **Deploy from GitHub**.
3. Pick the repo/branch and set **Main file path**:
   - `app.py` if this folder is the repo root, or
   - `eye-health-dashboard/app.py` if it sits in a sub-folder.
4. Click **Deploy** — `requirements.txt` is installed automatically.

> Streamlit Cloud looks for `requirements.txt` next to the app file or at the
> repo root. If you deploy with `eye-health-dashboard/` as a sub-folder and the
> build can't find the dependencies, copy `requirements.txt` to the repo root.

## Regenerate data / report (optional)

- `python prepare_data.py` rebuilds `data.js` and the static PNG charts from the
  source Excel file.
- `npm install && node build_report.js` rebuilds the Word data-story report.

## Notes

The dataset has **no date or geographic column**, so a date-range slider and map
are replaced by age / screen-time range sliders and a live Pearson correlation
heatmap. Eye-health score is a 0–100 composite. All correlations recompute with
the active filters.
