"""
Eye-Health Analytics — Streamlit deployment of the index.html dashboard.

Rather than re-implementing the UI with Streamlit widgets (which never matches
pixel-for-pixel), this serves the exact index.html — Plotly.js charts,
noUiSlider filters and all the cross-filtering — inside Streamlit via an
embedded component. data.js is inlined so no external file fetch is needed,
so it renders identically to opening index.html directly and deploys unchanged
on Streamlit Community Cloud.

Run locally:  streamlit run app.py
"""
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Eye-Health Analytics",
    page_icon="\U0001F441",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Strip all Streamlit chrome so the embedded dashboard fills the whole page and
# the page background matches index.html's paper colour (no white seams).
st.markdown(
    """
    <style>
      header[data-testid="stHeader"] { display: none; }
      [data-testid="stToolbar"], [data-testid="stDecoration"],
      [data-testid="stStatusWidget"], #MainMenu, footer { display: none; visibility: hidden; }
      [data-testid="stSidebarCollapsedControl"] { display: none; }
      .stApp { background: #eef4f3; }
      .block-container, [data-testid="stMainBlockContainer"] {
          padding: 0 !important; margin: 0 !important; max-width: 100% !important; }
      [data-testid="stAppViewContainer"] > .main { padding: 0 !important; }
      iframe { display: block; border: 0; }
    </style>
    """,
    unsafe_allow_html=True,
)

HERE = Path(__file__).parent
html = (HERE / "index.html").read_text(encoding="utf-8")
data_js = (HERE / "data.js").read_text(encoding="utf-8")

# Inline data.js in place of the external <script src="data.js"> so the embedded
# iframe (which can't resolve relative paths) still gets the dataset.
html = html.replace(
    '<script src="data.js"></script>',
    "<script>\n" + data_js + "\n</script>",
)

# Height matches index.html's rendered content (~1756px at desktop width) plus a
# little headroom; scrolling=True covers narrower viewports where it gets taller.
components.html(html, height=1820, scrolling=True)
