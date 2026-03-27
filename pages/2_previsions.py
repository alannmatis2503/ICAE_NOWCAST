"""Module 2 — Prévisions (mensuel / trimestriel / annuel)."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import CONSOLIDES, COUNTRY_CODES, COUNTRY_NAMES, CLOUD_MODE
from core.forecast_engine import (
    run_all_forecasts, get_methods_for_frequency, SEASON_MAP,
)
from core.tempdisagg import disaggregate_annual
from io_utils.excel_reader import (
    list_sheets, read_donnees_calcul, read_codification, rename_columns_to_codes,
)
from io_utils.excel_writer import write_previsions_excel, write_icae_recalc_output
from ui.charts import chart_forecast_comparison
from ui.components import download_button


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
    # Convertir en DatetimeIndex si c'est un Series (Series n'a pas .strftime)
    if isinstance(dates, pd.Series):
        dates = pd.DatetimeIndex(dates)
    if freq == "Mensuelle":
        return dates.strftime("%Y-%m")
    elif freq == "Trimestrielle":
        return [f"{d.year}T{(d.month - 1) // 3 + 1}" for d in dates]
    else:
        return dates.strftime("%Y")


def _get_active_vars(codif_df, available_cols):
    """Extrait les variables actives depuis la Codification (PRIOR > 0, Statut Actif)."""
    if codif_df is None:
        return None
    if "Code" not in codif_df.columns or "PRIOR" not in codif_df.columns:
        return None
    active_names = set()
    for _, row in codif_df.iterrows():
        try:
            if float(row.get("PRIOR", 0)) > 0:
                statut = str(row.get("Statut", "Actif")).lower()
                if not statut.startswith("inact"):
                    active_names.add(str(row["Code"]))
                    if "Label" in codif_df.columns and pd.notna(row.get("Label")):
                        active_names.add(str(row["Label"]))
        except (ValueError, TypeError):
            pass
    filtered = [c for c in available_cols if c in active_names]
    return filtered if filtered else None


st.title("📈 Module 2 — Prévisions")

# ── Mode ──────────────────────────────────────────────────────────────────
mode = st.radio("Mode", ["Prévision classique", "Désagrégation temporelle"],
                horizontal=True, key="prev_mode")

# ══════════════════════════════════════════════════════════════════════════
#  MODE DÉSAGRÉGATION TEMPORELLE
# ══════════════════════════════════════════════════════════════════════════
if mode == "Désagrégation temporelle":
    st.header("Désagrégation annuel → infra-annuel")
    st.info("Entrez une série annuelle et la méthode reproduira sa structure "
            "à la fréquence choisie (Chow-Lin ou Denton-Cholette).")

    target_freq = st.radio("Fréquence cible", ["Trimestrielle", "Mensuelle"],
                           horizontal=True, key="disagg_freq")

    st.subheader("1. Série annuelle")
    annual_file = st.file_uploader("Fichier Excel avec données annuelles",
                                   type=["xlsx"], key="disagg_annual_file")
    if annual_file is None:
        st.stop()

    annual_sheets = list_sheets(annual_file)
    annual_file.seek(0)
    annual_sheet = st.selectbox("Feuille", annual_sheets, key="disagg_annual_sheet")
    annual_file.seek(0)
    df_annual = pd.read_excel(annual_file, sheet_name=annual_sheet, engine="openpyxl")

    # Première colonne = année
    year_col = df_annual.columns[0]
    df_annual = df_annual.rename(columns={year_col: "Année"})
    df_annual["Année"] = df_annual["Année"].astype(str).str[:4].astype(int)
    df_annual = df_annual.set_index("Année")
    annual_vars = [c for c in df_annual.columns if df_annual[c].dtype in ("float64", "int64", "float32")]
    st.dataframe(df_annual[annual_vars].head(10), use_container_width=True)
    st.caption(f"{len(df_annual)} années × {len(annual_vars)} variables")

    # Indicateur HF optionnel
    st.subheader("2. Indicateur haute fréquence (optionnel)")
    st.caption("Si fourni, la méthode Chow-Lin sera utilisée. Sinon, Denton-Cholette.")
    hf_file = st.file_uploader("Fichier Excel HF (optionnel)", type=["xlsx"],
                                key="disagg_hf_file")
    hf_df = None
    if hf_file is not None:
        hf_sheets = list_sheets(hf_file)
        hf_file.seek(0)
        hf_sheet = st.selectbox("Feuille HF", hf_sheets, key="disagg_hf_sheet")
        hf_file.seek(0)
        hf_df = pd.read_excel(hf_file, sheet_name=hf_sheet, engine="openpyxl")
        # Supprimer la colonne date pour ne garder que les indicateurs
        hf_df = hf_df.select_dtypes(include="number")
        st.dataframe(hf_df.head(5), use_container_width=True)

    # Variables à désagréger
    selected_annual_vars = st.multiselect("Variables à désagréger", annual_vars,
                                          default=annual_vars[:5], key="disagg_vars")

    if st.button("🚀 Lancer la désagrégation", type="primary", key="run_disagg"):
        results_disagg = {}
        progress = st.progress(0)
        for i, var in enumerate(selected_annual_vars):
            series_a = df_annual[var].dropna()
            if len(series_a) < 3:
                continue
            result = disaggregate_annual(series_a, target_freq, hf_df)
            results_disagg[var] = result
            progress.progress((i + 1) / len(selected_annual_vars))
        st.session_state["disagg_results"] = results_disagg
        st.session_state["disagg_freq"] = target_freq
        method_used = "Chow-Lin" if hf_df is not None else "Denton-Cholette"
        st.success(f"✅ {len(results_disagg)} variables désagrégées ({method_used})")

    if "disagg_results" not in st.session_state:
        st.stop()

    results_disagg = st.session_state["disagg_results"]
    target_freq_r = st.session_state.get("disagg_freq", target_freq)

    st.header("3. Résultats")

    tab_table, tab_chart = st.tabs(["📊 Tableau", "📈 Graphique"])

    with tab_table:
        # Combiner toutes les séries désagrégées en un DataFrame
        if results_disagg:
            combined = pd.DataFrame(results_disagg)
            combined.index = combined.index.astype(str)
            st.dataframe(combined, use_container_width=True)

    with tab_chart:
        var_show = st.selectbox("Variable", list(results_disagg.keys()),
                                key="disagg_var_show")
        if var_show and var_show in results_disagg:
            import plotly.graph_objects as go
            series_hf = results_disagg[var_show]
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=[str(p) for p in series_hf.index],
                y=series_hf.values,
                mode="lines+markers", name=f"{var_show} ({target_freq_r})",
            ))
            # Superposer les valeurs annuelles
            if var_show in df_annual.columns:
                s_freq = 4 if target_freq_r == "Trimestrielle" else 12
                annual_x = [str(series_hf.index[i * s_freq]) for i in range(len(df_annual[var_show].dropna()))
                            if i * s_freq < len(series_hf)]
                annual_y = df_annual[var_show].dropna().values[:len(annual_x)]
                fig.add_trace(go.Bar(
                    x=annual_x, y=annual_y,
                    name=f"{var_show} (Annuel)", opacity=0.3,
                ))
            fig.update_layout(title=f"Désagrégation — {var_show}",
                              xaxis_title="Période", yaxis_title="Valeur")
            st.plotly_chart(fig, use_container_width=True)

    # Export
    st.header("4. Export")
    if st.button("📥 Exporter la désagrégation", key="export_disagg"):
        import io
        combined = pd.DataFrame(results_disagg)
        combined.index = combined.index.astype(str)
        combined.index.name = "Période"
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            combined.to_excel(writer, sheet_name="Désagrégation")
        download_button(
            buf.getvalue(),
            f"Desagregation_{target_freq_r[:4]}.xlsx",
            "📥 Télécharger",
        )

    st.stop()

# ══════════════════════════════════════════════════════════════════════════
#  MODE PRÉVISION CLASSIQUE (code existant)
# ══════════════════════════════════════════════════════════════════════════

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
codif_df = None
country = "XXX"

if source == "Données du Module 1" and has_session_data:
    available_countries = list(st.session_state["donnees_calcul"].keys())
    country = st.selectbox("Pays", available_countries, key="prev_country")
    donnees = st.session_state["donnees_calcul"][country]
    _codif_store = st.session_state.get("codification", {})
    if country in _codif_store:
        codif_df = _codif_store[country]
    st.session_state["prev_source_path"] = st.session_state.get("icae_filepath")
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

    # Lire la Codification si disponible dans le même classeur
    uploaded.seek(0)
    _up_sheets = list_sheets(uploaded)
    uploaded.seek(0)
    if "Codification" in _up_sheets:
        try:
            codif_df = read_codification(uploaded)
            uploaded.seek(0)
        except Exception:
            pass

    for c in COUNTRY_CODES:
        if c in uploaded.name.upper():
            country = c
            break

    st.session_state["prev_source_path"] = uploaded

    # Stocker la codification en session et renommer les colonnes si possible
    if codif_df is not None and country != "XXX":
        st.session_state.setdefault("codification", {})[country] = codif_df
        # Renommer les colonnes vers les codes si la codification est disponible
        try:
            donnees = rename_columns_to_codes(donnees, codif_df)
        except Exception:
            pass
elif source == "Fichier du dossier Livrables":
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
        if "Codification" in sheets:
            try:
                codif_df = read_codification(filepath)
            except Exception:
                pass
        for c in COUNTRY_CODES:
            if c in selected.upper():
                country = c
                break
        st.session_state["prev_source_path"] = filepath
        # Stocker la codification et renommer les colonnes
        if codif_df is not None and country != "XXX":
            st.session_state.setdefault("codification", {})[country] = codif_df
            try:
                donnees = rename_columns_to_codes(donnees, codif_df)
            except Exception:
                pass
    else:
        st.warning("Aucun fichier trouvé dans le dossier Livrables.")

if donnees is None:
    st.stop()

# ── Aperçu et période ────────────────────────────────────────────────────
data_cols = [c for c in donnees.columns if c != "Date"]
st.dataframe(donnees.head(10), use_container_width=True)
st.caption(f"{len(donnees)} observations × {len(data_cols)} variables")

# Sélection de la période
dates_series = pd.to_datetime(donnees["Date"], errors="coerce")
var_ranges = {}
for col in data_cols:
    valid_idx = donnees[col].notna() & dates_series.notna()
    valid_dates = dates_series[valid_idx]
    if len(valid_dates) > 0:
        var_ranges[col] = (valid_dates.min(), valid_dates.max())

if var_ranges:
    common_start = max(r[0] for r in var_ranges.values())
    common_end = min(r[1] for r in var_ranges.values())
    total_start = min(r[0] for r in var_ranges.values())
    total_end = max(r[1] for r in var_ranges.values())

    if common_start > common_end:
        st.warning("⚠️ Aucune période commune — période totale utilisée.")
        common_start, common_end = total_start, total_end

    with st.expander("📅 Période des données", expanded=True):
        st.caption(
            f"Période commune : **{common_start.strftime('%Y-%m')}** — "
            f"**{common_end.strftime('%Y-%m')}** | "
            f"Totale : {total_start.strftime('%Y-%m')} — {total_end.strftime('%Y-%m')}"
        )
        pc1, pc2 = st.columns(2)
        with pc1:
            start_date = st.date_input(
                "Début", value=common_start.date(),
                min_value=total_start.date(), max_value=total_end.date(),
                key="prev_start_date",
            )
        with pc2:
            end_date = st.date_input(
                "Fin", value=common_end.date(),
                min_value=total_start.date(), max_value=total_end.date(),
                key="prev_end_date",
            )
        mask = (dates_series >= pd.Timestamp(start_date)) & (
            dates_series <= pd.Timestamp(end_date)
        )
        donnees = donnees[mask].reset_index(drop=True)
        data_cols = [c for c in donnees.columns if c != "Date"]
        st.caption(f"✅ {len(donnees)} observations retenues")

if donnees.empty:
    st.warning("Aucune observation dans la période sélectionnée.")
    st.stop()

# ── Variables actives ────────────────────────────────────────────────────
active_vars = _get_active_vars(codif_df, data_cols)
has_codif = active_vars is not None

# ── Paramètres ────────────────────────────────────────────────────────────
st.header("2. Paramètres de prévision")

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
    var_options = (
        ["Variables actives (Codification)", "Toutes", "Sélection manuelle"]
        if has_codif
        else ["Toutes", "Sélection manuelle"]
    )
    var_mode = st.radio("Variables à prévoir", var_options, key="prev_var_mode")

if var_mode == "Variables actives (Codification)":
    selected_vars = active_vars
    st.info(f"🎯 {len(selected_vars)} variables actives depuis la Codification")
elif var_mode == "Sélection manuelle":
    default_sel = active_vars[:10] if has_codif else data_cols[:5]
    selected_vars = st.multiselect("Variables", data_cols, default=default_sel,
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
    st.session_state["prev_country_out"] = country
    st.session_state["prev_horizon_out"] = horizon
    st.session_state["prev_methods_out"] = methods
    st.session_state["prev_freq_out"] = freq
    st.session_state["prev_donnees_out"] = donnees
    st.success(f"✅ Prévisions calculées pour {len(all_results)} variables")

# ── Résultats interactifs ────────────────────────────────────────────────
if "prev_results" not in st.session_state:
    st.stop()

all_results = st.session_state["prev_results"]
horizon = st.session_state.get("prev_horizon_out", 3)
methods = st.session_state.get("prev_methods_out", available_methods[:6])
prev_freq = st.session_state.get("prev_freq_out", freq)
donnees = st.session_state.get("prev_donnees_out", donnees)

st.header("4. Résultats et sélection de méthode")

# ── Tableau comparatif MAPE ──────────────────────────────────────────────
tab_compare, tab_graph, tab_edit = st.tabs([
    "📊 Tableau comparatif", "📈 Graphiques par variable", "✏️ Édition manuelle",
])

with tab_compare:
    # Onglet de sélection MAPE / MAE / RMSE
    metric_choice = st.radio(
        "Critère de comparaison", ["MAPE", "MAE", "RMSE"],
        horizontal=True, key="prev_metric_choice",
    )
    metric_key = metric_choice.lower()

    st.subheader(f"{metric_choice} par méthode et variable")

    compare_rows = []
    for var, res in all_results.items():
        row = {"Variable": var}
        for m in methods:
            if m in res["backtesting"]:
                val = res["backtesting"][m].get(metric_key, np.nan)
                row[m] = round(val, 2) if not np.isnan(val) else None
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
            {"M\u00e9thode": m, "MAPE": round(res["backtesting"][m].get("mape", np.nan), 2),
             "MAE": round(res["backtesting"][m].get("mae", np.nan), 2),
             "RMSE": round(res["backtesting"][m].get("rmse", np.nan), 2)}
            for m in methods if m in res["backtesting"]
        ])
        st.dataframe(metrics_df, use_container_width=True, hide_index=True)

with tab_edit:
    st.subheader("Édition manuelle des prévisions")

    last_date = pd.to_datetime(donnees["Date"]).max()
    dates_fcst = _future_dates(last_date, horizon, prev_freq)

    # ── Contexte historique : 15 dernières périodes (lecture seule) ───────
    n_hist_ctx = min(15, len(donnees))
    _hist_vars = [v for v in all_results.keys() if v in donnees.columns]
    if _hist_vars:
        _hist_display = donnees.tail(n_hist_ctx)[["Date"] + _hist_vars].copy()
        _hist_display["Date"] = _format_dates(
            pd.to_datetime(_hist_display["Date"]), prev_freq
        )
        _hist_display = _hist_display.reset_index(drop=True)
        st.caption("📋 **Contexte historique** — 15 dernières périodes (lecture seule)")
        st.dataframe(_hist_display, use_container_width=True, hide_index=True)
    st.caption("✏️ **Prévisions à éditer** — modifiez les valeurs ci-dessous")

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

col1, col2, col3, col4 = st.columns(4)
with col1:
    if st.button("✅ Valider et recalculer l'ICAE", key="validate_prev",
                 type="primary"):
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

            # Auto-recalcul de l'ICAE en arrière-plan si fréquence mensuelle
            if prev_freq == "Mensuelle" and country != "XXX":
                try:
                    from core.icae_engine import run_icae_pipeline
                    from core.quarterly import (
                        quarterly_mean, calc_ga_trim, calc_gt_trim,
                        contributions_sectorielles_trim, normalize_contrib_to_ga,
                    )

                    _codif = st.session_state.get("codification", {}).get(country)
                    _base_year = st.session_state.get("icae_base_year_dict", {}).get(country, 2023)

                    dates_c = pd.to_datetime(extended["Date"])
                    base_mask = dates_c.dt.year == _base_year
                    base_indices = extended.index[base_mask]
                    if len(base_indices) > 0:
                        base_rows = range(base_indices[0], base_indices[-1] + 1)
                    else:
                        base_rows = range(108, 120)

                    data_cols_ext = [c for c in extended.columns if c != "Date"]
                    priors = pd.Series(1.0, index=data_cols_ext)
                    if _codif is not None and "Code" in _codif.columns and "PRIOR" in _codif.columns:
                        priors = pd.Series(
                            _codif["PRIOR"].values, index=_codif["Code"].values, dtype=float
                        ).fillna(0)

                    results_icae = run_icae_pipeline(
                        donnees=extended, priors=priors,
                        base_year=_base_year, base_rows=base_rows,
                    )
                    q = quarterly_mean(results_icae["icae"], results_icae["dates"])
                    q["GA_Trim"] = calc_ga_trim(q["icae_trim"])
                    q["GT_Trim"] = calc_gt_trim(q["icae_trim"])

                    contrib_trim = None
                    if "m" in results_icae["weights"]:
                        codif_for_contrib = _codif
                        # Si pas de codification, tenter de créer un pseudo-codif
                        if codif_for_contrib is None:
                            codif_for_contrib = pd.DataFrame({
                                "Code": results_icae["active_cols"],
                                "Secteur": "Autre",
                            })
                        contrib_trim = contributions_sectorielles_trim(
                            results_icae["weights"]["m"], codif_for_contrib,
                            results_icae["dates"],
                        )
                        if contrib_trim is not None:
                            contrib_trim = normalize_contrib_to_ga(
                                contrib_trim, q["GA_Trim"],
                            )

                    # Stocker tout en session state pour partage inter-modules
                    st.session_state["icae_results"] = results_icae
                    st.session_state["icae_country"] = country
                    st.session_state.setdefault("icae_monthly", {})[country] = results_icae["icae"]
                    st.session_state.setdefault("icae_quarterly", {})[country] = q
                    st.session_state.setdefault("icae_base_year_dict", {})[country] = _base_year
                    if contrib_trim is not None:
                        st.session_state.setdefault("icae_contrib_trim", {})[country] = contrib_trim

                    # Stocker le classeur complet recalculé pour partage
                    _recalc_wb = write_icae_recalc_output(
                        source_path=st.session_state.get("prev_source_path"),
                        donnees=extended,
                        results=results_icae,
                        quarterly=q,
                        contrib_trim=contrib_trim,
                        codification=_codif,
                        country_code=country,
                    )
                    st.session_state["_recalc_workbook"] = {country: _recalc_wb}

                    st.success("✅ Prévisions validées et ICAE recalculé automatiquement !")
                except Exception as e:
                    st.warning(f"Prévisions validées mais recalcul ICAE échoué : {e}")
                    import traceback
                    st.code(traceback.format_exc())
            else:
                st.success("✅ Prévisions validées")

with col2:
    # Bouton téléchargement du classeur ICAE recalculé
    _recalc_wb = st.session_state.get("_recalc_workbook", {}).get(country)
    if _recalc_wb is not None:
        download_button(
            _recalc_wb,
            f"ICAE_{country}_Recalcule_Series_Prolongees.xlsx",
            "📥 Télécharger l'ICAE recalculé",
        )
    else:
        st.button("📥 Télécharger l'ICAE recalculé", disabled=True,
                  key="download_recalc_disabled")

with col3:
    if st.button("📋 Aller au module Rapports", key="goto_rapports"):
        st.switch_page("pages/5_rapports.py")

with col4:
    if st.button("→ Nowcast avec séries prolongées", key="goto_nowcast"):
        st.session_state["_nowcast_use_extended"] = True
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
