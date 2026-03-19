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
    list_sheets, load_country_file, read_consignes, read_codification,
    read_donnees_calcul, rename_columns_to_codes,
)
from io_utils.excel_writer import write_cemac_excel
from ui.charts import chart_icae_monthly, chart_ga_bars
from ui.components import download_button


def _process_country_file(source):
    """Traite un fichier ICAE pays (chemin ou file-like) → (icae_series, dates) ou None."""
    try:
        if hasattr(source, 'seek'):
            source.seek(0)
        _sh = list_sheets(source)
        if hasattr(source, 'seek'):
            source.seek(0)
        if "Donnees_calcul" not in _sh:
            return None
        consignes = read_consignes(source) if "Consignes" in _sh else {"base_year": 2023}
        if hasattr(source, 'seek'):
            source.seek(0)
        codif = read_codification(source) if "Codification" in _sh else pd.DataFrame()
        if hasattr(source, 'seek'):
            source.seek(0)
        donnees = read_donnees_calcul(source)
        if not codif.empty:
            donnees = rename_columns_to_codes(donnees, codif)

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
        if not codif.empty and "Code" in codif.columns and "PRIOR" in codif.columns:
            priors = pd.Series(
                codif["PRIOR"].values, index=codif["Code"].values, dtype=float,
            ).fillna(0)
        else:
            data_cols = [c for c in donnees.columns if c != "Date"]
            priors = pd.Series(1.0, index=data_cols)

        results = run_icae_pipeline(
            donnees=donnees, priors=priors,
            base_year=base_year, base_rows=base_rows,
        )
        icae = results["icae"]
        icae.index = dates
        return icae, dates
    except Exception:
        return None


st.title("🌍 Module 4 — ICAE CEMAC Agrégé")

# ── Source ────────────────────────────────────────────────────────────────
st.header("1. Données")

has_icae = "icae_monthly" in st.session_state and len(st.session_state["icae_monthly"]) > 0

source_opts = []
if has_icae:
    source_opts.append("Données du Module 1")
source_opts.append("Upload de fichiers")
if not CLOUD_MODE and CONSOLIDES.exists():
    source_opts.append("Calcul depuis les fichiers consolidés")

source_mode = st.radio("Source", source_opts, horizontal=True, key="cemac_source")

icae_dict = {}
dates_dict = {}

if source_mode == "Données du Module 1" and has_icae:
    icae_dict = st.session_state["icae_monthly"]
    st.success(f"✅ ICAE disponibles pour : {', '.join(icae_dict.keys())}")

elif source_mode == "Upload de fichiers":
    st.info("Uploadez les fichiers consolidés ICAE pour chaque pays de la CEMAC.")
    uploaded_files = {}
    cols = st.columns(3)
    for idx, code in enumerate(COUNTRY_CODES):
        with cols[idx % 3]:
            f = st.file_uploader(
                f"{COUNTRY_NAMES[code]} ({code})",
                type=["xlsx"], key=f"cemac_upload_{code}",
            )
            if f is not None:
                uploaded_files[code] = f

    if not uploaded_files:
        st.warning("Veuillez uploader au moins un fichier.")
        st.stop()

    st.caption(f"📁 {len(uploaded_files)}/{len(COUNTRY_CODES)} fichiers uploadés")
    progress = st.progress(0)
    for i, (code, f) in enumerate(uploaded_files.items()):
        result = _process_country_file(f)
        if result is not None:
            icae_dict[code], dates_dict[code] = result
        else:
            st.warning(f"⚠️ Erreur pour {code}")
        progress.progress((i + 1) / len(uploaded_files))

    if icae_dict:
        st.session_state["icae_monthly"] = icae_dict
        st.success(f"✅ ICAE calculés pour : {', '.join(icae_dict.keys())}")

elif source_mode == "Calcul depuis les fichiers consolidés":
    st.info("Chargement des fichiers consolidés depuis le dossier Livrables...")
    progress = st.progress(0)
    for i, code in enumerate(COUNTRY_CODES):
        fpath = CONSOLIDES / f"ICAE_{code}_Consolide.xlsx"
        if not fpath.exists():
            st.warning(f"⚠️ Fichier manquant : {fpath.name}")
            progress.progress((i + 1) / len(COUNTRY_CODES))
            continue
        result = _process_country_file(fpath)
        if result is not None:
            icae_dict[code], dates_dict[code] = result
        else:
            st.warning(f"⚠️ Erreur pour {code}")
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
    st.session_state["cemac_computed_poids"] = filtered_poids
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
        st.session_state.get("cemac_computed_poids", POIDS_PIB),
    )
    download_button(
        data,
        "ICAE_CEMAC_Consolide_OUT.xlsx",
        "📥 Télécharger le fichier CEMAC",
    )
