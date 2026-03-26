"""Point d'entrée — Application ICAE / Nowcast / Prévisions."""
import sys
from pathlib import Path

# Ajouter le répertoire de l'app au path
APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import streamlit as st
from ui.styles import inject_css
from config import LOGO_PATH

st.set_page_config(
    page_title="Suivi du secteur productif infra-annuel : ICAE et PIB Nowcast",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_css()

# ── Sidebar ───────────────────────────────────────────────────────────────
with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=120)
    st.title("Suivi du secteur productif infra-annuel")
    st.caption("ICAE et PIB Nowcast — BEAC")
    st.markdown("---")

# ── Navigation ────────────────────────────────────────────────────────────
home_page = st.Page("pages/0_accueil.py", title="Accueil",
                    icon="🏠", default=True)
icae_page = st.Page("pages/1_icae.py", title="Module 1 — Calcul ICAE",
                    icon="📊")
prev_page = st.Page("pages/2_previsions.py", title="Module 2 — Prévisions",
                    icon="📈")
now_page = st.Page("pages/3_nowcast.py", title="Module 3 — Nowcast",
                   icon="🔮")
cemac_page = st.Page("pages/4_cemac.py", title="Module 4 — Agrégation CEMAC",
                     icon="🌍")
rap_page = st.Page("pages/5_rapports.py", title="Module 5 — Rapports",
                   icon="📋")
doc_page = st.Page("pages/6_documentation.py", title="Documentation / Aide",
                   icon="📄")

nav = st.navigation([home_page, icae_page, prev_page, now_page,
                      cemac_page, rap_page, doc_page])
nav.run()
