"""Module 4 — ICAE CEMAC agrégé."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import (CONSOLIDES, COUNTRY_CODES, COUNTRY_NAMES,
                    PIB_2014, POIDS_PIB, CLOUD_MODE)
from core.cemac_engine import compute_icae_cemac, quarterly_cemac
from core.icae_engine import run_icae_pipeline
from io_utils.excel_reader import (
    load_country_file, read_consignes, read_codification,
    read_donnees_calcul, rename_columns_to_codes,
)
from io_utils.excel_writer import write_cemac_excel
from ui.charts import chart_icae_monthly, chart_ga_bars
from ui.components import download_button

st.title("🌍 Module 4 — ICAE CEMAC Agrégé")

# ── Source ────────────────────────────────────────────────────────────────
st.header("1. Données")

has_icae = "icae_monthly" in st.session_state and len(st.session_state["icae_monthly"]) > 0

source_options = ["Données du Module 1"]
if not CLOUD_MODE:
    source_options.insert(0, "Calcul depuis les fichiers consolidés")
else:
    source_options.append("Upload des fichiers consolidés")

source_mode = st.radio(
    "Source",
    source_options,
    horizontal=True,
    key="cemac_source",
)

icae_dict = {}
dates_dict = {}

if source_mode == "Données du Module 1" and has_icae:
    icae_dict = st.session_state["icae_monthly"]
    st.success(f"✅ ICAE disponibles pour : {', '.join(icae_dict.keys())}")

elif source_mode == "Upload des fichiers consolidés":
    uploaded_files = st.file_uploader(
        "Fichiers ICAE consolidés (.xlsx) — un par pays",
        type=["xlsx"], accept_multiple_files=True,
        key="cemac_upload",
    )
    if uploaded_files:
        progress = st.progress(0)
        for i, ufile in enumerate(uploaded_files):
            # Extraire le code pays du nom de fichier (ICAE_CMR_Consolide.xlsx)
            fname = ufile.name
            code = None
            for c in COUNTRY_CODES:
                if c in fname.upper():
                    code = c
                    break
            if code is None:
                st.warning(f"⚠️ Code pays non reconnu dans : {fname}")
                continue
            try:
                consignes = read_consignes(ufile)
                codification = read_codification(ufile)
                donnees = read_donnees_calcul(ufile)
                donnees = rename_columns_to_codes(donnees, codification)

                base_year = consignes.get("base_year", 2023)
                dates = pd.to_datetime(donnees["Date"])
                base_mask = dates.dt.year == base_year
                base_indices = donnees.index[base_mask]

                if len(base_indices) > 0:
                    base_rows = range(base_indices[0], base_indices[-1] + 1)
                else:
                    bs = consignes.get("base_rows_start", 124) - 16
                    be = consignes.get("base_rows_end", 135) - 16
                    base_rows = range(bs, be + 1)

                priors = pd.Series(dtype=float)
                if "Code" in codification.columns and "PRIOR" in codification.columns:
                    priors = pd.Series(
                        codification["PRIOR"].values,
                        index=codification["Code"].values,
                        dtype=float,
                    ).fillna(0)
                else:
                    data_cols = [c for c in donnees.columns if c != "Date"]
                    priors = pd.Series(1.0, index=data_cols)

                results = run_icae_pipeline(
                    donnees=donnees, priors=priors,
                    base_year=base_year, base_rows=base_rows,
                )

                icae_dict[code] = results["icae"]
                icae_dict[code].index = dates
                dates_dict[code] = dates
            except Exception as e:
                st.warning(f"⚠️ Erreur pour {code} : {e}")

            progress.progress((i + 1) / len(uploaded_files))

        if icae_dict:
            st.session_state["icae_monthly"] = icae_dict
            st.success(f"✅ ICAE calculés pour : {', '.join(icae_dict.keys())}")
    else:
        st.info("Uploadez les 6 fichiers ICAE consolidés (un par pays CEMAC).")

else:
    st.info("Chargement des fichiers consolidés depuis le dossier Livrables...")
    
    if not CONSOLIDES.exists():
        st.error("Dossier Livrables introuvable.")
        st.stop()
    
    progress = st.progress(0)
    for i, code in enumerate(COUNTRY_CODES):
        fpath = CONSOLIDES / f"ICAE_{code}_Consolide.xlsx"
        if not fpath.exists():
            st.warning(f"⚠️ Fichier manquant : {fpath.name}")
            continue
        
        try:
            consignes = read_consignes(fpath)
            codification = read_codification(fpath)
            donnees = read_donnees_calcul(fpath)
            donnees = rename_columns_to_codes(donnees, codification)
            
            base_year = consignes.get("base_year", 2023)
            dates = pd.to_datetime(donnees["Date"])
            base_mask = dates.dt.year == base_year
            base_indices = donnees.index[base_mask]
            
            if len(base_indices) > 0:
                base_rows = range(base_indices[0], base_indices[-1] + 1)
            else:
                bs = consignes.get("base_rows_start", 124) - 16
                be = consignes.get("base_rows_end", 135) - 16
                base_rows = range(bs, be + 1)
            
            priors = pd.Series(dtype=float)
            if "Code" in codification.columns and "PRIOR" in codification.columns:
                priors = pd.Series(
                    codification["PRIOR"].values,
                    index=codification["Code"].values,
                    dtype=float,
                ).fillna(0)
            else:
                data_cols = [c for c in donnees.columns if c != "Date"]
                priors = pd.Series(1.0, index=data_cols)
            
            results = run_icae_pipeline(
                donnees=donnees,
                priors=priors,
                base_year=base_year,
                base_rows=base_rows,
            )
            
            icae_dict[code] = results["icae"]
            icae_dict[code].index = dates
            dates_dict[code] = dates
            
        except Exception as e:
            st.warning(f"⚠️ Erreur pour {code} : {e}")
        
        progress.progress((i + 1) / len(COUNTRY_CODES))
    
    if icae_dict:
        st.session_state["icae_monthly"] = icae_dict
        st.success(f"✅ ICAE calculés pour : {', '.join(icae_dict.keys())}")

if not icae_dict:
    st.warning("Aucun ICAE disponible.")
    st.stop()

# ── Pondérations ─────────────────────────────────────────────────────────
st.header("2. Pondérations PIB")

poids_data = []
for code in COUNTRY_CODES:
    poids_data.append({
        "Code": code,
        "Pays": COUNTRY_NAMES[code],
        "PIB 2014 (Mds FCFA)": PIB_2014[code],
        "Poids (%)": round(POIDS_PIB[code] * 100, 2),
    })

poids_df = pd.DataFrame(poids_data)
edited_poids = st.data_editor(poids_df, key="cemac_poids",
                              use_container_width=True, num_rows="fixed")

# Recalculer les poids
pib_edited = dict(zip(edited_poids["Code"], edited_poids["PIB 2014 (Mds FCFA)"]))
total_pib = sum(pib_edited.values())
poids = {code: pib_edited[code] / total_pib for code in pib_edited}

# Pays à inclure
pays_inclus = st.multiselect(
    "Pays inclus", COUNTRY_CODES,
    default=[c for c in COUNTRY_CODES if c in icae_dict],
    key="cemac_pays",
)

# ── Calcul CEMAC ─────────────────────────────────────────────────────────
st.header("3. Calcul ICAE CEMAC")

if st.button("🚀 Calculer l'ICAE CEMAC", type="primary", key="run_cemac"):
    filtered = {k: v for k, v in icae_dict.items() if k in pays_inclus}
    filtered_poids = {k: v for k, v in poids.items() if k in pays_inclus}
    
    # Renormaliser les poids
    total_w = sum(filtered_poids.values())
    filtered_poids = {k: v / total_w for k, v in filtered_poids.items()}
    
    result_df = compute_icae_cemac(filtered, filtered_poids)
    
    if result_df.empty:
        st.error("Pas assez de données pour calculer l'ICAE CEMAC.")
        st.stop()
    
    st.session_state["cemac_result"] = result_df
    st.session_state["cemac_poids"] = filtered_poids
    st.success("✅ ICAE CEMAC calculé !")

# ── Résultats ────────────────────────────────────────────────────────────
if "cemac_result" not in st.session_state:
    st.stop()

result_df = st.session_state["cemac_result"]

st.header("4. Résultats")

tab1, tab2, tab3 = st.tabs(["📈 ICAE mensuel", "📊 GA", "📋 Données"])

with tab1:
    fig = chart_icae_monthly(
        result_df.index, result_df["ICAE_CEMAC"],
        title="ICAE CEMAC — Agrégé pondéré",
    )
    # Ajouter les pays
    for code in pays_inclus:
        if code in result_df.columns:
            fig.add_scatter(x=result_df.index, y=result_df[code],
                           mode="lines", name=COUNTRY_NAMES.get(code, code),
                           opacity=0.5)
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    fig_ga = chart_ga_bars(
        result_df.index, result_df["GA"],
        title="ICAE CEMAC — Glissement annuel (%)",
    )
    st.plotly_chart(fig_ga, use_container_width=True)

with tab3:
    st.dataframe(result_df, use_container_width=True)

# Trimestriel
st.subheader("Résultats trimestriels")
dates_idx = pd.to_datetime(result_df.index)
q_cemac = quarterly_cemac(result_df, dates_idx)
st.dataframe(q_cemac, use_container_width=True, hide_index=True)

# ── Export ────────────────────────────────────────────────────────────────
st.header("5. Export")
if st.button("📥 Exporter CEMAC", key="export_cemac"):
    data = write_cemac_excel(
        result_df, q_cemac,
        st.session_state.get("cemac_poids", POIDS_PIB),
    )
    download_button(
        data,
        "ICAE_CEMAC_Consolide_OUT.xlsx",
        "📥 Télécharger le fichier CEMAC",
    )
