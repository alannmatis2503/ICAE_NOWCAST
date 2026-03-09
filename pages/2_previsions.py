"""Module 2 — Prévisions."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import CONSOLIDES, COUNTRY_CODES, COUNTRY_NAMES, FORECAST_METHODS
from core.forecast_engine import run_all_forecasts, METHOD_DISPATCH
from io_utils.excel_reader import list_sheets, read_donnees_calcul, read_codification
from io_utils.excel_writer import write_previsions_excel
from ui.charts import chart_forecast_comparison
from ui.components import download_button

st.title("📈 Module 2 — Prévisions")

# ── Source de données ─────────────────────────────────────────────────────
st.header("1. Données")
has_session_data = "donnees_calcul" in st.session_state and st.session_state["donnees_calcul"]

if has_session_data:
    st.success("✅ Données disponibles depuis le Module 1")
    available_countries = list(st.session_state["donnees_calcul"].keys())
    country = st.selectbox("Pays", available_countries, key="prev_country")
    donnees = st.session_state["donnees_calcul"][country]
else:
    st.info("Aucune donnée du Module 1. Veuillez charger un fichier.")
    uploaded = st.file_uploader("Fichier ICAE consolidé", type=["xlsx"],
                                key="prev_upload")
    if uploaded is None:
        st.stop()
    
    sheets = list_sheets(uploaded)
    uploaded.seek(0)
    sheet = st.selectbox("Feuille", sheets,
                         index=sheets.index("Donnees_calcul") if "Donnees_calcul" in sheets else 0,
                         key="prev_sheet")
    uploaded.seek(0)
    donnees = read_donnees_calcul(uploaded, sheet=sheet)
    country = "XXX"
    for c in COUNTRY_CODES:
        if c in uploaded.name:
            country = c
            break

# ── Paramètres ────────────────────────────────────────────────────────────
st.header("2. Paramètres de prévision")

data_cols = [c for c in donnees.columns if c != "Date"]

col1, col2 = st.columns(2)
with col1:
    horizon = st.slider("Horizon de prévision (mois)", 1, 24, 3,
                         key="prev_horizon")
    bt_window = st.slider("Fenêtre de backtesting (mois)", 6, 36, 12,
                           key="prev_bt_window")
with col2:
    methods = st.multiselect(
        "Méthodes de prévision", FORECAST_METHODS,
        default=["MM3", "MM6", "MM12", "NS", "CS", "TL"],
        key="prev_methods",
    )
    var_mode = st.radio("Variables à prévoir", ["Toutes", "Sélection manuelle"],
                        key="prev_var_mode")

if var_mode == "Sélection manuelle":
    selected_vars = st.multiselect("Variables", data_cols, default=data_cols[:5],
                                   key="prev_vars")
else:
    selected_vars = data_cols

if not methods:
    st.warning("Sélectionnez au moins une méthode.")
    st.stop()

# ── Calcul des prévisions ────────────────────────────────────────────────
st.header("3. Calcul")

if st.button("🚀 Lancer les prévisions", type="primary", key="run_prev"):
    progress = st.progress(0)
    all_results = {}
    
    for i, var in enumerate(selected_vars):
        series = donnees[var].dropna()
        if len(series) < 24:
            continue
        result = run_all_forecasts(series, horizon, methods, bt_window)
        all_results[var] = result
        progress.progress((i + 1) / len(selected_vars))

    st.session_state["prev_results"] = all_results
    st.session_state["prev_country"] = country
    st.session_state["prev_horizon"] = horizon
    st.session_state["prev_methods"] = methods
    st.success(f"✅ Prévisions calculées pour {len(all_results)} variables")

# ── Résultats interactifs ────────────────────────────────────────────────
if "prev_results" not in st.session_state:
    st.stop()

all_results = st.session_state["prev_results"]
horizon = st.session_state.get("prev_horizon", 3)
methods = st.session_state.get("prev_methods", FORECAST_METHODS[:6])

st.header("4. Résultats et sélection de méthode")

# ── Tableau comparatif MAPE ──────────────────────────────────────────────
tab_compare, tab_graph, tab_edit = st.tabs([
    "📊 Tableau comparatif", "📈 Graphiques par variable", "✏️ Édition manuelle",
])

with tab_compare:
    st.subheader("MAPE par méthode et variable")
    
    compare_rows = []
    for var, res in all_results.items():
        row = {"Variable": var}
        for m in methods:
            if m in res["backtesting"]:
                row[m] = round(res["backtesting"][m]["mape"], 2) \
                    if not np.isnan(res["backtesting"][m]["mape"]) else None
        row["Recommandée"] = res["best_method"]
        compare_rows.append(row)
    
    compare_df = pd.DataFrame(compare_rows)
    
    # Colorer le MAPE
    def _color_mape(val):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return "background-color: #f0f0f0"
        if val < 5:
            return "background-color: #d4edda; color: #155724"
        elif val < 10:
            return "background-color: #fff3cd; color: #856404"
        elif val < 20:
            return "background-color: #fde8d0; color: #c46210"
        else:
            return "background-color: #f8d7da; color: #721c24"
    
    method_cols = [c for c in compare_df.columns if c not in ("Variable", "Recommandée")]
    styled = compare_df.style.map(_color_mape, subset=method_cols)
    st.dataframe(styled, use_container_width=True, hide_index=True)

    # Permettre de changer la méthode recommandée
    st.subheader("Modifier la méthode recommandée")
    selected_methods = {}
    
    cols_per_row = 4
    var_list = list(all_results.keys())
    for i in range(0, len(var_list), cols_per_row):
        cols = st.columns(min(cols_per_row, len(var_list) - i))
        for j, col in enumerate(cols):
            if i + j < len(var_list):
                var = var_list[i + j]
                res = all_results[var]
                with col:
                    selected_methods[var] = st.selectbox(
                        var, methods,
                        index=methods.index(res["best_method"]) if res["best_method"] in methods else 0,
                        key=f"method_{var}",
                    )
    
    st.session_state["selected_methods"] = {country: selected_methods}

with tab_graph:
    st.subheader("Prévisions par variable")
    
    var_to_show = st.selectbox(
        "Variable à visualiser", list(all_results.keys()),
        key="prev_var_show",
    )
    
    if var_to_show in all_results:
        res = all_results[var_to_show]
        series = donnees[var_to_show].dropna()
        dates_hist = pd.to_datetime(donnees["Date"]).iloc[-min(36, len(series)):]
        series_show = series.iloc[-min(36, len(series)):]
        
        # Dates futures
        last_date = pd.to_datetime(donnees["Date"]).max()
        dates_fcst = pd.date_range(last_date + pd.DateOffset(months=1),
                                   periods=horizon, freq="MS")
        
        sel_method = selected_methods.get(var_to_show, res["best_method"])
        
        fig = chart_forecast_comparison(
            series_show.values, res["forecasts"], sel_method,
            dates_hist.values, dates_fcst, var_to_show,
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Tableau des métriques
        metrics_df = pd.DataFrame([
            {"Méthode": m, "MAPE": res["backtesting"][m]["mape"],
             "RMSE": res["backtesting"][m]["rmse"]}
            for m in methods if m in res["backtesting"]
        ])
        st.dataframe(metrics_df, use_container_width=True, hide_index=True)

with tab_edit:
    st.subheader("Édition manuelle des prévisions")
    
    last_date = pd.to_datetime(donnees["Date"]).max()
    dates_fcst = pd.date_range(last_date + pd.DateOffset(months=1),
                               periods=horizon, freq="MS")
    
    # Construire le tableau éditable
    edit_data = {"Date": dates_fcst.strftime("%Y-%m")}
    for var in all_results:
        sel = selected_methods.get(var, all_results[var]["best_method"])
        fc = all_results[var]["forecasts"].get(sel, np.full(horizon, np.nan))
        edit_data[var] = fc[:horizon] if len(fc) >= horizon else np.concatenate([fc, np.full(horizon - len(fc), np.nan)])
    
    edit_df = pd.DataFrame(edit_data)
    
    edited = st.data_editor(
        edit_df, use_container_width=True,
        key="prev_edit",
        num_rows="fixed",
    )
    
    st.session_state["forecasts_edited"] = {country: edited}

# ── Réinjection ──────────────────────────────────────────────────────────
st.header("5. Réinjection des prévisions")

col1, col2, col3 = st.columns(3)
with col1:
    if st.button("✅ Valider les prévisions", key="validate_prev"):
        edited = st.session_state.get("forecasts_edited", {}).get(country)
        if edited is not None:
            # Concaténer les prévisions aux données historiques
            last_date = pd.to_datetime(donnees["Date"]).max()
            dates_fcst = pd.date_range(last_date + pd.DateOffset(months=1),
                                       periods=horizon, freq="MS")
            new_rows = pd.DataFrame({"Date": dates_fcst})
            for col_name in edited.columns:
                if col_name != "Date" and col_name in donnees.columns:
                    new_rows[col_name] = edited[col_name].values
            
            extended = pd.concat([donnees, new_rows], ignore_index=True)
            st.session_state["donnees_calcul"][country] = extended
            st.session_state["forecasts"] = {country: edited}
            st.success("✅ Prévisions ajoutées aux séries")

with col2:
    if st.button("→ Recalculer l'ICAE", key="goto_icae"):
        st.switch_page("pages/1_icae.py")

with col3:
    if st.button("→ Nowcast avec séries prolongées", key="goto_nowcast"):
        st.switch_page("pages/3_nowcast.py")

# ── Export ────────────────────────────────────────────────────────────────
st.header("6. Export Excel")

if st.button("📥 Générer le fichier de prévisions", key="export_prev"):
    edited = st.session_state.get("forecasts_edited", {}).get(country)
    
    # Préparer les données d'export
    previsions = {}
    backtesting_data = {}
    recommandations = {}
    
    for var, res in all_results.items():
        previsions[var] = res["forecasts"]
        backtesting_data[var] = res["backtesting"]
        sel = selected_methods.get(var, res["best_method"])
        recommandations[var] = {
            "method": sel,
            "mape": res["backtesting"].get(sel, {}).get("mape", np.nan),
        }
    
    data = write_previsions_excel(
        donnees, previsions, backtesting_data, recommandations,
        edited=edited,
    )
    download_button(
        data,
        f"Prevision_ICAE_{country}_Q{(pd.Timestamp.now().month - 1) // 3 + 1}_{pd.Timestamp.now().year}.xlsx",
        "📥 Télécharger les prévisions",
    )
