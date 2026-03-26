"""Module 1 — Calcul ICAE."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import CONSOLIDES, COUNTRY_CODES, COUNTRY_NAMES
from core.icae_engine import run_icae_pipeline
from core.quarterly import (
    quarterly_mean, calc_ga_trim, calc_gt_trim,
    contributions_sectorielles, contributions_sectorielles_trim,
    normalize_contrib_to_ga,
)
from io_utils.excel_reader import (
    list_sheets, read_consignes, read_codification,
    read_donnees_calcul, load_country_file, rename_columns_to_codes,
)
from io_utils.excel_writer import write_icae_output
from ui.charts import (
    chart_icae_monthly, chart_ga_bars, chart_contributions,
    chart_quarterly_contrib_ga,
)
from ui.components import download_button

st.title("📊 Module 1 — Calcul ICAE")

# ── Source de données ─────────────────────────────────────────────────────
st.header("1. Chargement des données")

# Si des résultats existent déjà en session, permettre de les afficher
has_previous = "icae_results" in st.session_state

source_mode = st.radio(
    "Source des données",
    ["Upload d'un fichier", "Fichier du dossier Livrables"],
    horizontal=True,
    key="icae_source_mode",
)

filepath = None
if source_mode == "Upload d'un fichier":
    uploaded = st.file_uploader(
        "Fichier ICAE consolidé (.xlsx)", type=["xlsx"],
        key="icae_upload",
    )
    if uploaded:
        filepath = uploaded
else:
    available = []
    if CONSOLIDES.exists():
        available = sorted([
            f.name for f in CONSOLIDES.glob("ICAE_*_Consolide*.xlsx")
            if "CEMAC" not in f.name
        ])
    if available:
        selected = st.selectbox("Fichier", available, key="icae_file_select")
        filepath = CONSOLIDES / selected
    else:
        st.warning("Aucun fichier trouvé dans le dossier Livrables.")

if filepath is None and not has_previous:
    st.info("Veuillez charger ou sélectionner un fichier consolidé ICAE.")
    st.stop()

# ── Chargement et calcul (si fichier fourni) ─────────────────────────────
if filepath is not None:
    try:
        sheets = list_sheets(filepath)
        if hasattr(filepath, "seek"):
            filepath.seek(0)
    except Exception as e:
        st.error(f"Erreur à la lecture : {e}")
        st.stop()

    with st.expander("📋 Feuilles détectées", expanded=False):
        st.write(sheets)

    col1, col2 = st.columns(2)
    with col1:
        sheet_donnees = st.selectbox(
            "Feuille Données de calcul", sheets,
            index=sheets.index("Donnees_calcul") if "Donnees_calcul" in sheets else 0,
            key="sheet_donnees",
        )
    with col2:
        sheet_codif = st.selectbox(
            "Feuille Codification", sheets,
            index=sheets.index("Codification") if "Codification" in sheets else 0,
            key="sheet_codif",
        )

    @st.cache_data
    def _load_data(fp, sheet_d, sheet_c):
        if hasattr(fp, "read"):
            fp.seek(0)
        consignes = read_consignes(fp)
        if hasattr(fp, "seek"):
            fp.seek(0)
        codification = read_codification(fp, sheet=sheet_c)
        if hasattr(fp, "seek"):
            fp.seek(0)
        donnees = read_donnees_calcul(fp, sheet=sheet_d)
        return consignes, codification, donnees

    try:
        consignes, codification, donnees = _load_data(filepath, sheet_donnees, sheet_codif)
        donnees = rename_columns_to_codes(donnees, codification)
    except Exception as e:
        st.error(f"Erreur lors du chargement : {e}")
        st.stop()

    # ── Paramètres ────────────────────────────────────────────────────────
    st.header("2. Paramètres")

    country_code = "XXX"
    fname = filepath.name if hasattr(filepath, "name") else Path(filepath).name
    for c in COUNTRY_CODES:
        if c in fname:
            country_code = c
            break

    col1, col2, col3 = st.columns(3)
    with col1:
        base_year = st.number_input(
            "Année de base", min_value=2000, max_value=2030,
            value=consignes.get("base_year", 2023), key="base_year",
        )
    with col2:
        sigma_mode = st.radio("Écart-type", ["Fixe", "Glissant"], key="sigma_mode")
    with col3:
        rolling_window = 12
        if sigma_mode == "Glissant":
            rolling_window = st.slider("Fenêtre rolling", 6, 36, 12, key="rolling_window")

    # ── Période ───────────────────────────────────────────────────────────
    dates = pd.to_datetime(donnees["Date"], errors="coerce")
    data_cols_all = [c for c in donnees.columns if c != "Date"]
    last_valid = {}
    for c in data_cols_all:
        mask = donnees[c].notna() & (donnees[c] != "")
        if mask.any():
            last_valid[c] = dates[mask].max()

    full_start = dates.min()
    full_end = dates.max()
    common_end = min(last_valid.values()) if last_valid else full_end
    max_end = max(last_valid.values()) if last_valid else full_end

    st.subheader("Période de calcul")
    if max_end > common_end:
        st.info(
            f"ℹ️ Certaines variables disposent de données jusqu’au "
            f"**{max_end.strftime('%B %Y')}** tandis que d’autres s’arrêtent au "
            f"**{common_end.strftime('%B %Y')}**. Seules les variables actives "
            f"(PRIOR > 0) sont utilisées pour le calcul — vous pouvez étendre "
            f"la période sans risque."
        )
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        calc_start = st.date_input("Début", value=full_start.date(),
                                    min_value=full_start.date(),
                                    max_value=full_end.date(), key="calc_start")
    with col_p2:
        calc_end = st.date_input("Fin", value=max_end.date(),
                                  min_value=full_start.date(),
                                  max_value=full_end.date(), key="calc_end")

    period_mask = (dates >= pd.Timestamp(calc_start)) & (dates <= pd.Timestamp(calc_end))
    donnees = donnees.loc[period_mask].reset_index(drop=True)

    with st.expander("📊 Valeurs par défaut détectées", expanded=True):
        defaults_df = pd.DataFrame([
            {"Paramètre": "Année de base", "Valeur": consignes.get("base_year", "?")},
            {"Paramètre": "Pays détecté", "Valeur": f"{country_code} — {COUNTRY_NAMES.get(country_code, '?')}"},
            {"Paramètre": "Nombre de variables", "Valeur": len(donnees.columns) - 1},
            {"Paramètre": "Période de calcul", "Valeur": f"{calc_start} → {calc_end} ({len(donnees)} obs.)"},
        ])
        st.dataframe(defaults_df, use_container_width=True, hide_index=True)

    # ── Codification ──────────────────────────────────────────────────────
    st.header("3. Variables et pondérations")

    if "Code" in codification.columns and "PRIOR" in codification.columns:
        codif_edit = codification.copy()
        if "Statut" not in codif_edit.columns:
            codif_edit["Statut"] = "Actif"

        edited_codif = st.data_editor(codif_edit, use_container_width=True,
                                      key="codif_editor", num_rows="fixed")

        priors = pd.Series(edited_codif["PRIOR"].values,
                           index=edited_codif["Code"].values, dtype=float).fillna(0)
        active_mask = (priors > 0)
        if "Statut" in edited_codif.columns:
            statut_map = dict(zip(edited_codif["Code"], edited_codif["Statut"]))
            for code in priors.index:
                if str(statut_map.get(code, "Actif")).lower().startswith("inact"):
                    active_mask[code] = False
        st.info(f"**{active_mask.sum()} variables actives** sur {len(priors)}")
    else:
        st.warning("Format de Codification non reconnu. Utilisation de poids égaux.")
        data_cols = [c for c in donnees.columns if c != "Date"]
        priors = pd.Series(1.0, index=data_cols)

    # ── Calcul ICAE ──────────────────────────────────────────────────────
    st.header("4. Calcul")

    if st.button("🚀 Lancer le calcul ICAE", type="primary", key="run_icae"):
        dates_calc = pd.to_datetime(donnees["Date"])
        base_mask = dates_calc.dt.year == base_year
        base_indices = donnees.index[base_mask]
        if len(base_indices) > 0:
            base_rows = range(base_indices[0], base_indices[-1] + 1)
        else:
            base_start = consignes.get("base_rows_start", 124)
            base_end_c = consignes.get("base_rows_end", 135)
            row_offset = 16
            base_rows = range(base_start - row_offset, base_end_c - row_offset + 1)

        with st.spinner("Calcul en cours..."):
            try:
                results = run_icae_pipeline(
                    donnees=donnees, priors=priors, base_year=base_year,
                    base_rows=base_rows,
                    sigma_mode="fixed" if sigma_mode == "Fixe" else "rolling",
                    rolling_window=rolling_window,
                )
                # Calculer les résultats trimestriels
                q = quarterly_mean(results["icae"], results["dates"])
                q["GA_Trim"] = calc_ga_trim(q["icae_trim"])
                q["GT_Trim"] = calc_gt_trim(q["icae_trim"])

                # Contributions sectorielles trimestrielles
                codif_for_contrib = edited_codif if "edited_codif" in dir() else codification
                contrib_trim = None
                if "m" in results["weights"]:
                    contrib_trim = contributions_sectorielles_trim(
                        results["weights"]["m"], codif_for_contrib, results["dates"],
                    )
                    # Normaliser pour que Σ(contributions) = GA à chaque trimestre
                    if contrib_trim is not None:
                        contrib_trim = normalize_contrib_to_ga(
                            contrib_trim, q["GA_Trim"],
                        )

                # Détecter la frontière historique / prévision
                all_vars = [c for c in donnees.columns if c != "Date"]
                active_set = set(results["active_cols"])
                non_active_vars = [c for c in all_vars if c not in active_set]
                fcst_boundary_q = None
                if non_active_vars:
                    dates_det = pd.to_datetime(donnees["Date"])
                    na_ends = []
                    for c in non_active_vars:
                        vmask = donnees[c].notna() & (donnees[c].astype(str).str.strip() != "")
                        if vmask.any():
                            na_ends.append(dates_det[vmask].max())
                    if na_ends:
                        boundary = min(na_ends)
                        active_ends = []
                        for c in results["active_cols"]:
                            vmask = donnees[c].notna() & (donnees[c].astype(str).str.strip() != "")
                            if vmask.any():
                                active_ends.append(dates_det[vmask].max())
                        if active_ends and max(active_ends) > boundary:
                            next_m = boundary + pd.DateOffset(months=1)
                            fcst_boundary_q = (
                                f"{next_m.year}T{(next_m.month - 1) // 3 + 1}"
                            )
                st.session_state.setdefault(
                    "icae_forecast_boundary", {},
                )[country_code] = fcst_boundary_q

                # Stocker tout en session_state
                st.session_state["icae_results"] = results
                st.session_state["icae_country"] = country_code
                st.session_state["donnees_calcul"] = st.session_state.get("donnees_calcul", {})
                st.session_state["donnees_calcul"][country_code] = donnees
                st.session_state["codification"] = st.session_state.get("codification", {})
                st.session_state["codification"][country_code] = codif_for_contrib
                st.session_state["icae_monthly"] = st.session_state.get("icae_monthly", {})
                st.session_state["icae_monthly"][country_code] = results["icae"]
                st.session_state["icae_base_year_dict"] = st.session_state.get("icae_base_year_dict", {})
                st.session_state["icae_base_year_dict"][country_code] = base_year
                st.session_state["icae_quarterly"] = st.session_state.get("icae_quarterly", {})
                st.session_state["icae_quarterly"][country_code] = q
                st.session_state["icae_contrib_trim"] = st.session_state.get("icae_contrib_trim", {})
                if contrib_trim is not None:
                    st.session_state["icae_contrib_trim"][country_code] = contrib_trim
                st.session_state["icae_filepath"] = filepath
                st.success("✅ Calcul terminé !")
            except Exception as e:
                st.error(f"❌ Erreur de calcul : {e}")
                import traceback
                st.code(traceback.format_exc())
                st.stop()

# ── Affichage des résultats (depuis session_state) ───────────────────────
if "icae_results" not in st.session_state:
    st.stop()

results = st.session_state["icae_results"]
country_code = st.session_state.get("icae_country", "XXX")
_base_year = st.session_state.get("icae_base_year_dict", {}).get(country_code, 2023)

st.header("5. Résultats")

# ── Filtre de période pour les graphiques (slider start/end) ──────────────
_all_dates = pd.to_datetime(results["dates"])
_date_min, _date_max = _all_dates.min(), _all_dates.max()

_col_s1, _col_s2 = st.columns(2)
with _col_s1:
    _chart_start = pd.Timestamp(st.date_input(
        "📅 Début d'affichage", value=_date_min.date(),
        min_value=_date_min.date(), max_value=_date_max.date(),
        key="icae_chart_start"))
with _col_s2:
    _chart_end = pd.Timestamp(st.date_input(
        "📅 Fin d'affichage", value=_date_max.date(),
        min_value=_date_min.date(), max_value=_date_max.date(),
        key="icae_chart_end"))

_chart_mask = (_all_dates >= _chart_start) & (_all_dates <= _chart_end)
_chart_dates = _all_dates[_chart_mask]
_chart_icae = np.array(results["icae"])[_chart_mask]
_chart_ga = np.array(results["ga_monthly"])[_chart_mask]

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 ICAE mensuel", "📊 GA mensuel",
    "📊 Trimestriel (GA + Contributions)", "🏗️ Contributions mensuelles",
    "📋 Données",
])

with tab1:
    fig = chart_icae_monthly(
        _chart_dates, _chart_icae,
        title=f"ICAE {country_code} — Base {_base_year} = 100",
    )
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    fig_ga = chart_ga_bars(
        _chart_dates, _chart_ga,
        title=f"ICAE {country_code} — Glissement annuel mensuel (%)",
    )
    st.plotly_chart(fig_ga, use_container_width=True)

with tab3:
    # Graphique trimestriel : barres empilées contrib + courbe GA
    q_data = st.session_state.get("icae_quarterly", {}).get(country_code)
    ct_data = st.session_state.get("icae_contrib_trim", {}).get(country_code)
    if q_data is not None and ct_data is not None:
        # Filtrer les trimestres selon la période sélectionnée
        _q_trim = q_data["trimestre"].tolist()
        _q_filter_mask = []
        for t in _q_trim:
            try:
                p = pd.Period(t, freq="Q")
                _q_filter_mask.append(p.start_time >= _chart_start and p.end_time <= _chart_end + pd.Timedelta(days=31))
            except Exception:
                _q_filter_mask.append(True)
        _q_filtered = q_data[_q_filter_mask].reset_index(drop=True)
        _ct_filtered = ct_data.iloc[:len(q_data)][_q_filter_mask].reset_index(drop=True) if len(ct_data) >= len(q_data) else ct_data

        trimestres = _q_filtered["trimestre"].tolist()
        ga_trim = _q_filtered["GA_Trim"]
        fig_qt = chart_quarterly_contrib_ga(
            _ct_filtered, ga_trim, trimestres,
            title=f"ICAE {country_code} — Contributions sectorielles et GA trimestriel (%)",
        )
        st.plotly_chart(fig_qt, use_container_width=True)
    elif q_data is not None:
        import plotly.graph_objects as go
        fig_gt = go.Figure()
        fig_gt.add_trace(go.Scatter(
            x=q_data["trimestre"], y=q_data["GA_Trim"],
            mode="lines+markers", name="GA ICAE (%)",
            line=dict(color="#C00000", width=3),
        ))
        fig_gt.update_layout(title=f"GA trimestriel ICAE {country_code}",
                             template="plotly_white")
        st.plotly_chart(fig_gt, use_container_width=True)

    if q_data is not None:
        st.dataframe(q_data, use_container_width=True, hide_index=True)

with tab4:
    if "m" in results["weights"]:
        codif_for = st.session_state.get("codification", {}).get(
            country_code,
            codification if "codification" in dir() else pd.DataFrame())
        contrib = contributions_sectorielles(
            results["weights"]["m"], codif_for, results["dates"],
        )
        # Appliquer le filtre de période aux contributions mensuelles
        _contrib_dates = pd.to_datetime(results["dates"])
        _contrib_mask = (_contrib_dates >= _chart_start) & (_contrib_dates <= _chart_end)
        _contrib_filtered = contrib[_contrib_mask.values].reset_index(drop=True)
        _contrib_dates_filtered = _contrib_dates[_contrib_mask]

        fig_c = chart_contributions(
            _contrib_dates_filtered, _contrib_filtered,
            title=f"Contributions sectorielles mensuelles — {country_code}",
        )
        st.plotly_chart(fig_c, use_container_width=True)

with tab5:
    res_df = pd.DataFrame({
        "Date": results["dates"],
        "ICAE": results["icae"],
        "Indice": results["indice"],
        "Σm": results["sum_m"],
        "GA (%)": results["ga_monthly"],
    })
    st.dataframe(res_df, use_container_width=True, hide_index=True)

# ── Pondérations ──────────────────────────────────────────────────────────
with st.expander("⚖️ Pondérations finales"):
    pond = results["pond_finale"]
    if isinstance(pond, pd.Series):
        st.bar_chart(pond)
    elif isinstance(pond, pd.DataFrame):
        st.dataframe(pond.tail(12), use_container_width=True)

# ── Export ────────────────────────────────────────────────────────────────
st.header("6. Export")
_fp = st.session_state.get("icae_filepath", filepath if "filepath" in dir() else None)
if _fp is not None and st.button("📥 Générer le fichier Excel", key="export_icae"):
    try:
        q_data = st.session_state.get("icae_quarterly", {}).get(country_code)
        q_export = q_data.copy() if q_data is not None else pd.DataFrame()
        for col in ("quarter", "debut", "fin"):
            if col in q_export.columns:
                q_export[col] = q_export[col].astype(str)
        export_results = {
            "icae": results["icae"],
            "indice": results["indice"],
            "sum_m": results["sum_m"],
            "pond_finale": results["pond_finale"],
            "quarterly": q_export,
            "dates": results["dates"],
        }
        if hasattr(_fp, "seek"):
            _fp.seek(0)
        data = write_icae_output(_fp, export_results, country_code)
        download_button(data, f"ICAE_{country_code}_Consolide_OUT.xlsx",
                        "📥 Télécharger le fichier ICAE")
    except Exception as e:
        st.error(f"Erreur d'export : {e}")
