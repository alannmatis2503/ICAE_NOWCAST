"""Module 2 — Prévisions (mensuel / trimestriel / annuel)."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import CONSOLIDES, COUNTRY_CODES, COUNTRY_NAMES, CLOUD_MODE
from core.forecast_engine import (
    run_all_forecasts, get_methods_for_frequency, SEASON_MAP,
)
from io_utils.excel_reader import list_sheets, read_donnees_calcul, read_codification
from io_utils.excel_writer import write_previsions_excel
from ui.charts import chart_forecast_comparison
from ui.components import download_button

st.title("📈 Module 2 — Prévisions")

# ── Fréquence ────────────────────────────────────────────────────────────
st.header("1. Fréquence et données")

freq = st.radio("Fréquence des séries", ["Mensuelle", "Trimestrielle", "Annuelle"],
                horizontal=True, key="prev_freq")

# ── Source de données ─────────────────────────────────────────────────────
has_session_data = (freq == "Mensuelle"
                    and "donnees_calcul" in st.session_state
                    and st.session_state["donnees_calcul"])

source_opts = ["Upload d'un fichier"]
if has_session_data:
    source_opts.insert(0, "Données du Module 1")
if not CLOUD_MODE:
    source_opts.append("Fichier du dossier Livrables")

source = st.radio("Source", source_opts, horizontal=True, key="prev_source")

donnees = None
country = "XXX"

if source == "Données du Module 1" and has_session_data:
    available_countries = list(st.session_state["donnees_calcul"].keys())
    country = st.selectbox("Pays", available_countries, key="prev_country")
    donnees = st.session_state["donnees_calcul"][country]
elif source == "Upload d'un fichier":
    uploaded = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"],
                                key="prev_upload")
    if uploaded is None:
        st.stop()
    sheets = list_sheets(uploaded)
    uploaded.seek(0)

    # Détection heuristique de la feuille par défaut
    default_idx = 0
    freq_keywords = {
        "Mensuelle": ["donnees_calcul", "mensuel", "monthly", "data"],
        "Trimestrielle": ["trim", "quarter", "pib_trim", "trimestr"],
        "Annuelle": ["annuel", "annual", "yearly", "pib_annuel"],
    }
    for i, sn in enumerate(sheets):
        sn_lower = sn.lower()
        for kw in freq_keywords.get(freq, []):
            if kw in sn_lower:
                default_idx = i
                break

    sheet = st.selectbox("Feuille", sheets, index=default_idx, key="prev_sheet")
    uploaded.seek(0)

    # Lire les données de manière générique
    df = pd.read_excel(uploaded, sheet_name=sheet, engine="openpyxl")
    date_col = df.columns[0]
    df = df.rename(columns={date_col: "Date"})

    # Parser les dates selon la fréquence
    if freq == "Trimestrielle":
        df["Date"] = df["Date"].apply(_parse_quarterly_date)
    elif freq == "Annuelle":
        df["Date"] = pd.to_datetime(df["Date"].astype(str).str[:4], format="%Y",
                                    errors="coerce")
    else:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    donnees = df

    for c in COUNTRY_CODES:
        if c in uploaded.name.upper():
            country = c
            break
else:
    # Fichier du dossier Livrables
    available = []
    if CONSOLIDES.exists():
        available = sorted([
            f.name for f in CONSOLIDES.glob("ICAE_*_Consolide.xlsx")
            if "CEMAC" not in f.name
        ])
    if available:
        selected = st.selectbox("Fichier", available, key="prev_file_select")
        filepath = CONSOLIDES / selected
        sheets = list_sheets(filepath)
        sheet = st.selectbox("Feuille", sheets,
                             index=sheets.index("Donnees_calcul")
                             if "Donnees_calcul" in sheets else 0,
                             key="prev_sheet_loc")
        donnees = read_donnees_calcul(filepath, sheet=sheet)
        for c in COUNTRY_CODES:
            if c in selected.upper():
                country = c
                break
    else:
        st.warning("Aucun fichier trouvé dans le dossier Livrables.")

if donnees is None:
    st.stop()

# Afficher un aperçu
st.dataframe(donnees.head(10), use_container_width=True)
st.caption(f"{len(donnees)} observations × {len(donnees.columns) - 1} variables")

# ── Paramètres ────────────────────────────────────────────────────────────
st.header("2. Paramètres de prévision")

data_cols = [c for c in donnees.columns if c != "Date"]
available_methods = get_methods_for_frequency(freq)

freq_labels = {"Mensuelle": "mois", "Trimestrielle": "trimestres", "Annuelle": "années"}
freq_unit = freq_labels[freq]

col1, col2 = st.columns(2)
with col1:
    h_max = {"Mensuelle": 24, "Trimestrielle": 12, "Annuelle": 5}[freq]
    h_default = {"Mensuelle": 3, "Trimestrielle": 4, "Annuelle": 2}[freq]
    horizon = st.slider(f"Horizon de prévision ({freq_unit})", 1, h_max, h_default,
                         key="prev_horizon")
    bt_max = {"Mensuelle": 36, "Trimestrielle": 16, "Annuelle": 5}[freq]
    bt_default = {"Mensuelle": 12, "Trimestrielle": 8, "Annuelle": 3}[freq]
    bt_window = st.slider(f"Fenêtre de backtesting ({freq_unit})", 2, bt_max, bt_default,
                           key="prev_bt_window")
with col2:
    default_methods = [m for m in available_methods if m not in ("ARIMA", "ETS")]
    methods = st.multiselect(
        "Méthodes de prévision", available_methods,
        default=default_methods,
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

min_obs = {"Mensuelle": 24, "Trimestrielle": 8, "Annuelle": 4}[freq]

if st.button("🚀 Lancer les prévisions", type="primary", key="run_prev"):
    progress = st.progress(0)
    all_results = {}

    for i, var in enumerate(selected_vars):
        series = donnees[var].dropna()
        if len(series) < min_obs:
            continue
        result = run_all_forecasts(series, horizon, methods, bt_window, freq=freq)
        all_results[var] = result
        progress.progress((i + 1) / len(selected_vars))

    st.session_state["prev_results"] = all_results
    st.session_state["prev_country"] = country
    st.session_state["prev_horizon"] = horizon
    st.session_state["prev_methods"] = methods
    st.session_state["prev_freq"] = freq
    st.session_state["prev_donnees"] = donnees
    st.success(f"✅ Prévisions calculées pour {len(all_results)} variables")

# ── Résultats interactifs ────────────────────────────────────────────────
if "prev_results" not in st.session_state:
    st.stop()

all_results = st.session_state["prev_results"]
horizon = st.session_state.get("prev_horizon", 3)
methods = st.session_state.get("prev_methods", available_methods[:6])
prev_freq = st.session_state.get("prev_freq", freq)
donnees = st.session_state.get("prev_donnees", donnees)

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

    # Modifier la méthode recommandée
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
        show_n = min({"Mensuelle": 36, "Trimestrielle": 16, "Annuelle": 10}[prev_freq],
                     len(series))
        dates_hist = pd.to_datetime(donnees["Date"]).iloc[-show_n:]
        series_show = series.iloc[-show_n:]

        # Dates futures selon la fréquence
        last_date = pd.to_datetime(donnees["Date"]).max()
        dates_fcst = _future_dates(last_date, horizon, prev_freq)

        sel_method = selected_methods.get(var_to_show, res["best_method"])

        fig = chart_forecast_comparison(
            series_show.values, res["forecasts"], sel_method,
            dates_hist.values, dates_fcst, var_to_show,
        )
        st.plotly_chart(fig, use_container_width=True)

        metrics_df = pd.DataFrame([
            {"Méthode": m, "MAPE": res["backtesting"][m]["mape"],
             "RMSE": res["backtesting"][m]["rmse"]}
            for m in methods if m in res["backtesting"]
        ])
        st.dataframe(metrics_df, use_container_width=True, hide_index=True)

with tab_edit:
    st.subheader("Édition manuelle des prévisions")

    last_date = pd.to_datetime(donnees["Date"]).max()
    dates_fcst = _future_dates(last_date, horizon, prev_freq)

    # Construire le tableau éditable
    edit_data = {"Date": _format_dates(dates_fcst, prev_freq)}
    for var in all_results:
        sel = selected_methods.get(var, all_results[var]["best_method"])
        fc = all_results[var]["forecasts"].get(sel, np.full(horizon, np.nan))
        edit_data[var] = fc[:horizon] if len(fc) >= horizon else np.concatenate(
            [fc, np.full(horizon - len(fc), np.nan)])

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
            last_date = pd.to_datetime(donnees["Date"]).max()
            dates_fcst = _future_dates(last_date, horizon, prev_freq)
            new_rows = pd.DataFrame({"Date": dates_fcst})
            for col_name in edited.columns:
                if col_name != "Date" and col_name in donnees.columns:
                    new_rows[col_name] = edited[col_name].values

            extended = pd.concat([donnees, new_rows], ignore_index=True)
            if "donnees_calcul" not in st.session_state:
                st.session_state["donnees_calcul"] = {}
            st.session_state["donnees_calcul"][country] = extended
            st.session_state["forecasts"] = {country: edited}
            st.success("✅ Prévisions ajoutées aux séries")

with col2:
    if prev_freq == "Mensuelle":
        if st.button("→ Recalculer l'ICAE", key="goto_icae"):
            st.switch_page("pages/1_icae.py")

with col3:
    if st.button("→ Nowcast avec séries prolongées", key="goto_nowcast"):
        st.switch_page("pages/3_nowcast.py")

# ── Export ────────────────────────────────────────────────────────────────
st.header("6. Export Excel")

if st.button("📥 Générer le fichier de prévisions", key="export_prev"):
    edited = st.session_state.get("forecasts_edited", {}).get(country)

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

    freq_tag = {"Mensuelle": "M", "Trimestrielle": "T", "Annuelle": "A"}[prev_freq]
    download_button(
        data,
        f"Prevision_{country}_{freq_tag}_Q{(pd.Timestamp.now().month - 1) // 3 + 1}_{pd.Timestamp.now().year}.xlsx",
        "📥 Télécharger les prévisions",
    )


# ────────────────────────────────────────────────────────────────────────────
# Fonctions utilitaires (parsing dates multi-fréquence)
# ────────────────────────────────────────────────────────────────────────────
def _parse_quarterly_date(val):
    """Parse '2024T1', '2024Q1', '2024-Q1' etc. en Timestamp."""
    s = str(val).strip().upper()
    for sep in ["T", "Q", "-Q", "-T"]:
        if sep in s:
            parts = s.split(sep[-1])
            if len(parts) == 2:
                try:
                    year, q = int(parts[0].replace("-", "")), int(parts[1])
                    month = (q - 1) * 3 + 1
                    return pd.Timestamp(year=year, month=month, day=1)
                except (ValueError, TypeError):
                    pass
    return pd.NaT


def _future_dates(last_date, horizon, freq):
    """Génère les dates futures selon la fréquence."""
    last = pd.Timestamp(last_date)
    if freq == "Mensuelle":
        return pd.date_range(last + pd.DateOffset(months=1),
                             periods=horizon, freq="MS")
    elif freq == "Trimestrielle":
        return pd.date_range(last + pd.DateOffset(months=3),
                             periods=horizon, freq="QS")
    else:  # Annuelle
        return pd.date_range(last + pd.DateOffset(years=1),
                             periods=horizon, freq="YS")


def _format_dates(dates, freq):
    """Formate les dates pour l'affichage selon la fréquence."""
    if freq == "Mensuelle":
        return dates.strftime("%Y-%m")
    elif freq == "Trimestrielle":
        return [f"{d.year}T{(d.month - 1) // 3 + 1}" for d in dates]
    else:
        return dates.strftime("%Y")
