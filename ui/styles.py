"""Styles CSS personnalisés."""

CUSTOM_CSS = """
<style>
    /* En-tête */
    .main .block-container { padding-top: 1rem; }
    
    /* Tableaux */
    .stDataFrame { font-size: 0.85rem; }
    
    /* Couleurs historique / prévision */
    .hist-cell { background-color: #DCE6F1; color: #1F4E79; }
    .fcst-cell { background-color: #FDE9D9; color: #E26B0A; font-weight: bold; }
    
    /* Métriques */
    .metric-good { color: #28a745; font-weight: bold; }
    .metric-ok { color: #ffc107; font-weight: bold; }
    .metric-bad { color: #dc3545; font-weight: bold; }
    
    /* Sidebar */
    [data-testid="stSidebar"] { min-width: 280px; }
    
    /* Titres */
    h1 { color: #1F4E79; }
    h2 { color: #2E75B6; border-bottom: 2px solid #2E75B6; padding-bottom: 5px; }
</style>
"""


def inject_css():
    """Injecte le CSS personnalisé dans la page."""
    import streamlit as st
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
