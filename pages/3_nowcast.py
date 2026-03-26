"""Module 3 — Nowcast."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import COUNTRY_CODES, COUNTRY_NAMES, NOWCAST_MODELS
from core.nowcast_engine import run_nowcast
from core.quarterly import agg_m_to_q
from core.tempdisagg import disaggregate_annual_to_quarterly
from io_utils.excel_reader import list_sheets
from io_utils.excel_writer import write_nowcast_excel
from ui.charts import chart_nowcast, chart_ga_nowcast
from ui.components import download_button

st.title("🔮 Module 3 — Nowcast")

# ── Données HF ───────────────────────────────────────────────────────────
st.header("1. Indicateurs haute fréquence")

hf_source = st.radio(
    "Source des indicateurs HF",
    ["Upload d'un fichier", "Données du Module 1"],
    horizontal=True,
    key="now_hf_source",
)

hf_df = None
if hf_source == "Données du Module 1" and "donnees_calcul" in st.session_state:
    available = list(st.session_state["donnees_calcul"].keys())
    country = st.selectbox("Pays", available, key="now_country")
    hf_df = st.session_state["donnees_calcul"][country]
    st.success(f"✅ Données HF chargées depuis Module 1 ({country})")
else:
    uploaded_hf = st.file_uploader("Fichier des indicateurs mensuels", type=["xlsx"],
                                    key="now_hf_upload")
    if uploaded_hf:
        sheets = list_sheets(uploaded_hf)
        uploaded_hf.seek(0)
        sheet_hf = st.selectbox("Feuille HF", sheets, key="now_hf_sheet")
        uploaded_hf.seek(0)
        hf_df = pd.read_excel(uploaded_hf, sheet_name=sheet_hf, engine="openpyxl")
        country = "XXX"
        for c in COUNTRY_CODES:
            if c in uploaded_hf.name:
                country = c
                break

if hf_df is None:
    st.info("Veuillez charger les indicateurs haute fréquence.")
    st.stop()

# Sélection date et variables HF
hf_cols = list(hf_df.columns)
col1, col2 = st.columns(2)
with col1:
    date_col = st.selectbox("Colonne date", hf_cols,
                            index=0, key="now_date_col")
with col2:
    non_date = [c for c in hf_cols if c != date_col]
    # Par défaut : variables actives de la Codification (PRIOR > 0, Statut Actif)
    default_hf = non_date
    codif_dict = st.session_state.get("codification", {})
    if codif_dict:
        # Chercher la codification du pays courant
        codif_df = next(iter(codif_dict.values()), None)
        if codif_df is not None and "Code" in codif_df.columns and "PRIOR" in codif_df.columns:
            active_codes = set()
            for _, row in codif_df.iterrows():
                try:
                    if float(row.get("PRIOR", 0)) > 0:
                        statut = str(row.get("Statut", "Actif")).lower()
                        if not statut.startswith("inact"):
                            active_codes.add(row["Code"])
                            if "Label" in codif_df.columns:
                                active_codes.add(row["Label"])
                except (ValueError, TypeError):
                    pass
            filtered = [c for c in non_date if c in active_codes]
            if filtered:
                default_hf = filtered
    hf_vars = st.multiselect("Variables HF", non_date, default=default_hf,
                             key="now_hf_vars")

if not hf_vars:
    st.warning("Sélectionnez au moins une variable HF.")
    st.stop()

# Types d'agrégation
with st.expander("⚙️ Types d'agrégation mensuel→trimestriel"):
    agg_types = {}
    cols_per_row = 4
    for i in range(0, len(hf_vars), cols_per_row):
        cols = st.columns(min(cols_per_row, len(hf_vars) - i))
        for j, col in enumerate(cols):
            if i + j < len(hf_vars):
                v = hf_vars[i + j]
                with col:
                    # Détection automatique par préfixe
                    default = "Moyenne"
                    vl = v.lower()
                    if vl.startswith("st_") or vl.startswith("stock"):
                        default = "Stock (dernier)"
                    elif vl.startswith("flu_") or vl.startswith("flux"):
                        default = "Flux (somme)"
                    agg_types[v] = st.selectbox(
                        v[:20], ["Moyenne", "Stock (dernier)", "Flux (somme)"],
                        index=["Moyenne", "Stock (dernier)", "Flux (somme)"].index(default),
                        key=f"agg_{v}",
                    )
    agg_map = {v: {"Moyenne": "mean", "Stock (dernier)": "stock",
                   "Flux (somme)": "flow"}[t] for v, t in agg_types.items()}

# ── PIB ───────────────────────────────────────────────────────────────────
st.header("2. PIB")

uploaded_pib = st.file_uploader("Fichier du PIB", type=["xlsx"],
                                key="now_pib_upload")
pib_q = None

if uploaded_pib:
    sheets_pib = list_sheets(uploaded_pib)
    uploaded_pib.seek(0)
    sheet_pib = st.selectbox("Feuille PIB", sheets_pib, key="now_pib_sheet")
    uploaded_pib.seek(0)
    pib_raw = pd.read_excel(uploaded_pib, sheet_name=sheet_pib, engine="openpyxl")
    
    pib_cols = list(pib_raw.columns)
    col1, col2 = st.columns(2)
    with col1:
        pib_date_col = st.selectbox("Colonne date/année", pib_cols,
                                    key="now_pib_date")
    with col2:
        pib_val_col = st.selectbox("Colonne PIB", [c for c in pib_cols if c != pib_date_col],
                                   key="now_pib_val")
    
    pib_series = pib_raw[[pib_date_col, pib_val_col]].dropna()
    
    # Détection fréquence
    sample = pib_series[pib_date_col].iloc[:5]
    is_annual = all(isinstance(v, (int, float)) and v > 1900 and v < 2100
                    for v in sample)
    
    if is_annual:
        st.info("📅 PIB détecté comme **annuel**.")
        auto_disagg = st.checkbox("Trimestrialiser automatiquement (Chow-Lin)",
                                  key="now_auto_disagg")
        
        pib_annual = pd.Series(
            pib_series[pib_val_col].values,
            index=pib_series[pib_date_col].astype(int).values,
        )
        
        if auto_disagg:
            with st.spinner("Trimestrialisation en cours..."):
                # Construire les HF trimestriels
                hf_data = hf_df[[date_col] + hf_vars].copy()
                hf_data[date_col] = pd.to_datetime(hf_data[date_col], errors="coerce")
                hf_data = hf_data.dropna(subset=[date_col])
                hf_num = hf_data[hf_vars].apply(pd.to_numeric, errors="coerce")
                hf_q_data = agg_m_to_q(hf_num, hf_data[date_col], agg_map)
                
                if "quarter" in hf_q_data.columns:
                    hf_q_for_td = hf_q_data.drop(columns=["quarter"])
                else:
                    hf_q_for_td = hf_q_data
                
                pib_q = disaggregate_annual_to_quarterly(pib_annual, hf_q_for_td)
                st.success("✅ PIB trimestrialisé avec succès")
        else:
            # Distribution uniforme simple
            pib_q = disaggregate_annual_to_quarterly(pib_annual)
            st.info("PIB distribué uniformément (sans indicateur)")
    else:
        # PIB trimestriel
        pib_series[pib_date_col] = pd.to_datetime(pib_series[pib_date_col])
        pib_q = pd.Series(
            pib_series[pib_val_col].values,
            index=pib_series[pib_date_col].dt.to_period("Q"),
        )
        st.success("📅 PIB détecté comme **trimestriel**.")

if pib_q is None:
    st.info("Veuillez charger le fichier du PIB.")
    st.stop()

# ── Paramètres Nowcast ───────────────────────────────────────────────────
st.header("3. Paramètres")

col1, col2, col3 = st.columns(3)
with col1:
    models = st.multiselect("Modèles", NOWCAST_MODELS, default=NOWCAST_MODELS,
                            key="now_models")
with col2:
    n_components = st.slider("Nb facteurs (PC/DFM)", 1, 5, 2,
                             key="now_ncomp")
with col3:
    h_ahead = st.slider("Horizon (trimestres)", 1, 8, 4,
                         key="now_h_ahead")

# ── Calcul ────────────────────────────────────────────────────────────────
st.header("4. Calcul")

if st.button("🚀 Lancer le Nowcast", type="primary", key="run_nowcast"):
    with st.spinner("Modèles en cours d'estimation..."):
        # Agréger HF en trimestriel
        hf_data = hf_df[[date_col] + hf_vars].copy()
        hf_data[date_col] = pd.to_datetime(hf_data[date_col], errors="coerce")
        hf_data = hf_data.dropna(subset=[date_col])
        hf_num = hf_data[hf_vars].apply(pd.to_numeric, errors="coerce")
        hf_q = agg_m_to_q(hf_num, hf_data[date_col], agg_map)
        
        # Index par quarter
        if "quarter" in hf_q.columns:
            hf_q = hf_q.set_index("quarter")
        
        # Aligner PIB et HF
        pib_aligned = pib_q.copy()
        if hasattr(pib_aligned.index, "to_period"):
            pass
        elif not isinstance(pib_aligned.index, pd.PeriodIndex):
            pib_aligned.index = pd.PeriodIndex(pib_aligned.index, freq="Q")
        
        results = run_nowcast(
            pib_aligned, hf_q,
            models=models,
            h_ahead=h_ahead,
            n_components=n_components,
        )
        
        st.session_state["nowcast_results"] = {country: results}
        st.session_state["nowcast_pib"] = {country: pib_aligned}
        st.success("✅ Nowcast terminé !")

# ── Résultats ────────────────────────────────────────────────────────────
if "nowcast_results" not in st.session_state:
    st.stop()

results = list(st.session_state["nowcast_results"].values())[0]
pib_aligned = list(st.session_state["nowcast_pib"].values())[0]

st.header("5. Résultats")

tab1, tab2, tab3 = st.tabs(["📈 Graphiques", "📊 Performance", "📉 GA Nowcast"])

with tab1:
    fig = chart_nowcast(pib_aligned, results,
                        title=f"PIB observé vs Nowcasts — {country}")
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    perf_rows = []
    for name, r in results.items():
        m = r["metrics"]
        perf_rows.append({
            "Modèle": name,
            "RMSE (in)": round(m["in_sample"].get("rmse", np.nan), 2),
            "MAE (in)": round(m["in_sample"].get("mae", np.nan), 2),
            "MAPE (in)": round(m["in_sample"].get("mape", np.nan), 2),
            "RMSE (out)": round(m["out_sample"].get("rmse", np.nan), 2),
            "MAE (out)": round(m["out_sample"].get("mae", np.nan), 2),
            "MAPE (out)": round(m["out_sample"].get("mape", np.nan), 2),
            "Corrélation": round(r.get("correlation", np.nan), 4),
        })
    st.dataframe(pd.DataFrame(perf_rows), use_container_width=True, hide_index=True)

with tab3:
    fig_ga = chart_ga_nowcast(pib_aligned, results,
                               title=f"GA du PIB et des Nowcasts — {country}")
    st.plotly_chart(fig_ga, use_container_width=True)

# ── Export ────────────────────────────────────────────────────────────────
st.header("6. Export")
if st.button("📥 Exporter les résultats Nowcast", key="export_nowcast"):
    params = {
        "Pays": country,
        "Modèles": ", ".join(models),
        "Nb facteurs": n_components,
        "Horizon": h_ahead,
    }
    data = write_nowcast_excel(pib_aligned, results, params)
    download_button(
        data,
        f"RESULT_NOWCAST_{country}_{pd.Timestamp.now().strftime('%Y-%m-%d')}.xlsx",
        "📥 Télécharger le Nowcast",
    )
