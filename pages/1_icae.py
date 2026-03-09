"""Module 1 — Calcul ICAE."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import CONSOLIDES, COUNTRY_CODES, COUNTRY_NAMES, CLOUD_MODE
from core.icae_engine import run_icae_pipeline
from core.quarterly import quarterly_mean, calc_ga_trim, calc_gt_trim, contributions_sectorielles
from io_utils.excel_reader import (
    list_sheets, read_consignes, read_codification,
    read_donnees_calcul, load_country_file, rename_columns_to_codes,
)
from io_utils.excel_writer import write_icae_output
from ui.charts import chart_icae_monthly, chart_ga_bars, chart_contributions
from ui.components import download_button

st.title("📊 Module 1 — Calcul ICAE")

# ── Source de données ─────────────────────────────────────────────────────
st.header("1. Chargement des données")
source_options = ["Upload d'un fichier"]
if not CLOUD_MODE:
    source_options.append("Fichier du dossier Livrables")
source_mode = st.radio(
    "Source des données",
    source_options,
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
            f.name for f in CONSOLIDES.glob("ICAE_*_Consolide.xlsx")
            if "CEMAC" not in f.name
        ])
    if available:
        selected = st.selectbox("Fichier", available, key="icae_file_select")
        filepath = CONSOLIDES / selected
    else:
        st.warning("Aucun fichier trouvé dans le dossier Livrables.")

if filepath is None:
    st.info("Veuillez charger ou sélectionner un fichier consolidé ICAE.")
    st.stop()

# ── Sélection de feuilles ────────────────────────────────────────────────
try:
    sheets = list_sheets(filepath)
    if hasattr(filepath, "seek"):
        filepath.seek(0)
except Exception as e:
    st.error(f"Erreur à la lecture : {e}")
    st.stop()

with st.expander("📋 Feuilles détectées", expanded=False):
    st.write(sheets)

# Sélection des feuilles
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

# ── Lecture des données ──────────────────────────────────────────────────
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
    # Renommer colonnes Label→Code pour aligner avec priors
    donnees = rename_columns_to_codes(donnees, codification)
except Exception as e:
    st.error(f"Erreur lors du chargement : {e}")
    st.stop()

# ── Paramètres détectés ──────────────────────────────────────────────────
st.header("2. Paramètres")

# Déterminer le code pays
country_code = "XXX"
if hasattr(filepath, "name"):
    fname = filepath.name
else:
    fname = Path(filepath).name
for c in COUNTRY_CODES:
    if c in fname:
        country_code = c
        break

col1, col2, col3 = st.columns(3)
with col1:
    base_year = st.number_input(
        "Année de base",
        min_value=2000, max_value=2030,
        value=consignes.get("base_year", 2023),
        key="base_year",
    )
with col2:
    sigma_mode = st.radio(
        "Mode écart-type",
        ["Fixe", "Glissant"],
        key="sigma_mode",
    )
with col3:
    rolling_window = 12
    if sigma_mode == "Glissant":
        rolling_window = st.slider("Fenêtre rolling", 6, 36, 12,
                                   key="rolling_window")

# Afficher les valeurs par défaut
with st.expander("📊 Valeurs par défaut détectées", expanded=True):
    defaults_df = pd.DataFrame([
        {"Paramètre": "Année de base", "Valeur": consignes.get("base_year", "?")},
        {"Paramètre": "Lignes base (STDEV)", "Valeur": f"{consignes.get('base_rows_start', '?')}:{consignes.get('base_rows_end', '?')}"},
        {"Paramètre": "Pays détecté", "Valeur": f"{country_code} — {COUNTRY_NAMES.get(country_code, '?')}"},
        {"Paramètre": "Nombre de variables", "Valeur": len(donnees.columns) - 1},
        {"Paramètre": "Période", "Valeur": f"{donnees['Date'].min()} → {donnees['Date'].max()}"},
        {"Paramètre": "Facteur de symétrie", "Valeur": "200 (C = 200×(X−X₋₁)/(X+X₋₁))"},
    ])
    st.dataframe(defaults_df, use_container_width=True, hide_index=True)

# ── Codification et statuts ──────────────────────────────────────────────
st.header("3. Variables et pondérations")

if "Code" in codification.columns and "PRIOR" in codification.columns:
    # Éditable
    codif_edit = codification.copy()
    if "Statut" not in codif_edit.columns:
        codif_edit["Statut"] = "Actif"

    edited_codif = st.data_editor(
        codif_edit,
        use_container_width=True,
        key="codif_editor",
        num_rows="fixed",
    )

    # Extraire les priors
    priors = pd.Series(
        edited_codif["PRIOR"].values,
        index=edited_codif["Code"].values,
        dtype=float,
    ).fillna(0)

    # Variables actives = PRIOR > 0 ET Statut == Actif
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

# ── Calcul ICAE ──────────────────────────────────────────────────────────
st.header("4. Calcul")

if st.button("🚀 Lancer le calcul ICAE", type="primary", key="run_icae"):
    # Calculer les lignes de base dans le TCS
    base_start = consignes.get("base_rows_start", 124)
    base_end = consignes.get("base_rows_end", 135)
    # Convertir les lignes Excel (1-based, relative à row 16) en index Python
    # Row 16 dans Excel = index 0 dans le TCS
    # Les lignes base sont relatives aux données dans CALCUL_ICAE
    n_data = len(donnees)
    dates = pd.to_datetime(donnees["Date"])
    base_mask = dates.dt.year == base_year
    base_indices = donnees.index[base_mask]
    if len(base_indices) > 0:
        base_rows = range(base_indices[0], base_indices[-1] + 1)
    else:
        # Fallback : utiliser les lignes du fichier (ajustées)
        row_offset = 16  # données commencent à la ligne 16 dans Excel
        base_rows = range(base_start - row_offset, base_end - row_offset + 1)

    with st.spinner("Calcul en cours..."):
        try:
            results = run_icae_pipeline(
                donnees=donnees,
                priors=priors,
                base_year=base_year,
                base_rows=base_rows,
                sigma_mode="fixed" if sigma_mode == "Fixe" else "rolling",
                rolling_window=rolling_window,
            )
            st.session_state["icae_results"] = results
            st.session_state["icae_country"] = country_code
            st.session_state["donnees_calcul"] = {country_code: donnees}
            st.session_state["codification"] = {country_code: codification if "codification" not in dir() else edited_codif}
            st.session_state["icae_monthly"] = {country_code: results["icae"]}
            st.session_state["base_year"] = {country_code: base_year}
            st.success("✅ Calcul terminé !")
        except Exception as e:
            st.error(f"❌ Erreur de calcul : {e}")
            import traceback
            st.code(traceback.format_exc())
            st.stop()

# ── Affichage des résultats ──────────────────────────────────────────────
if "icae_results" in st.session_state:
    results = st.session_state["icae_results"]
    st.header("5. Résultats")

    # Onglets
    tab1, tab2, tab3, tab4 = st.tabs([
        "📈 ICAE mensuel", "📊 Glissement annuel",
        "🏗️ Contributions", "📋 Données",
    ])

    with tab1:
        fig = chart_icae_monthly(
            results["dates"], results["icae"],
            title=f"ICAE {country_code} — Base {base_year} = 100",
        )
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        ga = results["ga_monthly"]
        fig_ga = chart_ga_bars(
            results["dates"], ga,
            title=f"ICAE {country_code} — Glissement annuel (%)",
        )
        st.plotly_chart(fig_ga, use_container_width=True)

    with tab3:
        if "m" in results["weights"]:
            codif_for_contrib = edited_codif if "edited_codif" in dir() else codification
            contrib = contributions_sectorielles(
                results["weights"]["m"], codif_for_contrib, results["dates"],
            )
            fig_c = chart_contributions(
                results["dates"], contrib,
                title=f"Contributions sectorielles — {country_code}",
            )
            st.plotly_chart(fig_c, use_container_width=True)

    with tab4:
        # Tableau des résultats
        res_df = pd.DataFrame({
            "Date": results["dates"],
            "ICAE": results["icae"],
            "Indice": results["indice"],
            "Σm": results["sum_m"],
            "GA (%)": results["ga_monthly"],
        })
        st.dataframe(res_df, use_container_width=True, hide_index=True)

    # Trimestriel
    st.subheader("Résultats trimestriels")
    q = quarterly_mean(results["icae"], results["dates"])
    q["GA_Trim"] = calc_ga_trim(q["icae_trim"])
    q["GT_Trim"] = calc_gt_trim(q["icae_trim"])
    st.dataframe(q, use_container_width=True, hide_index=True)

    # Stocker pour les autres modules
    st.session_state["icae_quarterly"] = {country_code: q}

    # Pondérations
    with st.expander("⚖️ Pondérations finales"):
        pond = results["pond_finale"]
        if isinstance(pond, pd.Series):
            st.bar_chart(pond)
        elif isinstance(pond, pd.DataFrame):
            st.dataframe(pond.tail(12), use_container_width=True)

    # ── Export ────────────────────────────────────────────────────────────
    st.header("6. Export")
    if st.button("📥 Générer le fichier Excel", key="export_icae"):
        try:
            # Préparer les résultats pour l'export
            q_export = q.copy()
            if "quarter" in q_export.columns:
                q_export["quarter"] = q_export["quarter"].astype(str)
            if "debut" in q_export.columns:
                q_export["debut"] = q_export["debut"].astype(str)
            if "fin" in q_export.columns:
                q_export["fin"] = q_export["fin"].astype(str)
            export_results = {
                "icae": results["icae"],
                "indice": results["indice"],
                "sum_m": results["sum_m"],
                "pond_finale": results["pond_finale"],
                "quarterly": q_export,
            }
            if hasattr(filepath, "seek"):
                filepath.seek(0)
            data = write_icae_output(filepath, export_results, country_code)
            download_button(
                data,
                f"ICAE_{country_code}_Consolide_OUT.xlsx",
                "📥 Télécharger le fichier ICAE",
            )
        except Exception as e:
            st.error(f"Erreur d'export : {e}")
