"""Module 4 — ICAE CEMAC agrégé."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import (CONSOLIDES, COUNTRY_CODES, COUNTRY_NAMES,
                    PIB_2014, POIDS_PIB, CLOUD_MODE, CEMAC_TEMPLATE)
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

# ── Période commune ──────────────────────────────────────────────────────
st.header("2. Période de calcul")

# Déterminer la fenêtre temporelle commune des séries pays
_country_ranges = {}
for code, series in icae_dict.items():
    if hasattr(series, "index"):
        idx = pd.to_datetime(series.index)
        valid = idx[series.notna().values] if hasattr(series, "notna") else idx
        if len(valid) > 0:
            _country_ranges[code] = (valid.min(), valid.max())

if _country_ranges:
    _common_start = max(r[0] for r in _country_ranges.values())
    _common_end = min(r[1] for r in _country_ranges.values())
    _total_start = min(r[0] for r in _country_ranges.values())
    _total_end = max(r[1] for r in _country_ranges.values())

    with st.expander("📅 Fenêtre temporelle des séries pays", expanded=True):
        _range_rows = []
        for code, (s, e) in _country_ranges.items():
            _range_rows.append({
                "Pays": COUNTRY_NAMES.get(code, code),
                "Début": s.strftime("%Y-%m"),
                "Fin": e.strftime("%Y-%m"),
            })
        st.dataframe(pd.DataFrame(_range_rows), use_container_width=True,
                     hide_index=True)
        st.caption(
            f"Période commune : **{_common_start.strftime('%Y-%m')}** — "
            f"**{_common_end.strftime('%Y-%m')}** | "
            f"Totale : {_total_start.strftime('%Y-%m')} — {_total_end.strftime('%Y-%m')}"
        )

    _pc1, _pc2 = st.columns(2)
    with _pc1:
        _cemac_start = pd.Timestamp(st.date_input(
            "Début", value=_common_start.date(),
            min_value=_total_start.date(), max_value=_total_end.date(),
            key="cemac_start_date",
        ))
    with _pc2:
        _cemac_end = pd.Timestamp(st.date_input(
            "Fin", value=_common_end.date(),
            min_value=_total_start.date(), max_value=_total_end.date(),
            key="cemac_end_date",
        ))

    # Filtrer les séries selon la période choisie
    icae_dict_filtered = {}
    for code, series in icae_dict.items():
        if hasattr(series, "index"):
            idx = pd.to_datetime(series.index)
            mask = (idx >= _cemac_start) & (idx <= _cemac_end)
            icae_dict_filtered[code] = series[mask]
    icae_dict = icae_dict_filtered
else:
    _total_start = None
    _total_end = None

# ── Pondérations ─────────────────────────────────────────────────────────
st.header("3. Pondérations PIB")

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
st.header("4. Calcul ICAE CEMAC")

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

st.header("5. Résultats")

# ── Fenêtre temporelle d'affichage des graphiques ────────────────────────
st.subheader("📅 Fenêtre d'affichage des graphiques")
_res_dates = pd.to_datetime(result_df.index)
_res_min, _res_max = _res_dates.min(), _res_dates.max()

_gc1, _gc2 = st.columns(2)
with _gc1:
    _graph_start = pd.Timestamp(st.date_input(
        "Début d'affichage", value=_res_min.date(),
        min_value=_res_min.date(), max_value=_res_max.date(),
        key="cemac_graph_start",
    ))
with _gc2:
    _graph_end = pd.Timestamp(st.date_input(
        "Fin d'affichage", value=_res_max.date(),
        min_value=_res_min.date(), max_value=_res_max.date(),
        key="cemac_graph_end",
    ))

_graph_mask = (_res_dates >= _graph_start) & (_res_dates <= _graph_end)
_disp_df = result_df[_graph_mask]

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 ICAE mensuel", "📊 GA",
    "📊 Contributions pays", "📋 Taux de croissance (8 trim.)",
    "📋 Données",
])

with tab1:
    fig = chart_icae_monthly(
        _disp_df.index, _disp_df["ICAE_CEMAC"],
        title="ICAE CEMAC — Agrégé pondéré",
    )
    # Ajouter les pays
    for code in COUNTRY_CODES:
        if code in _disp_df.columns and code in pays_inclus:
            fig.add_scatter(x=_disp_df.index, y=_disp_df[code],
                           mode="lines", name=COUNTRY_NAMES.get(code, code),
                           opacity=0.5)
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    fig_ga = chart_ga_bars(
        _disp_df.index, _disp_df["GA"],
        title="ICAE CEMAC — Glissement annuel (%)",
    )
    st.plotly_chart(fig_ga, use_container_width=True)

with tab3:
    # Graphique des contributions pays à la croissance CEMAC
    import plotly.graph_objects as go
    _poids_used = st.session_state.get("cemac_computed_poids", POIDS_PIB)
    _country_colors = {
        "CMR": "#4472C4", "RCA": "#ED7D31", "CNG": "#A5A5A5",
        "GAB": "#FFC000", "GNQ": "#5B9BD5", "TCD": "#70AD47",
    }
    fig_contrib = go.Figure()
    for code in COUNTRY_CODES:
        if code in _disp_df.columns and code in pays_inclus:
            # Contribution pays = poids * GA_pays
            ga_pays = (_disp_df[code] / _disp_df[code].shift(12) - 1) * 100
            contrib_pays = ga_pays * _poids_used.get(code, 0)
            fig_contrib.add_trace(go.Bar(
                x=_disp_df.index, y=contrib_pays,
                name=f"{COUNTRY_NAMES.get(code, code)} ({round(_poids_used.get(code, 0) * 100, 1)}%)",
                marker_color=_country_colors.get(code, "#888"),
            ))
    # Courbe GA CEMAC
    fig_contrib.add_trace(go.Scatter(
        x=_disp_df.index, y=_disp_df["GA"],
        mode="lines+markers", name="GA CEMAC (%)",
        line=dict(color="#C00000", width=3),
    ))
    fig_contrib.update_layout(
        barmode="relative",
        title="Contributions pays au GA de l'ICAE CEMAC (%)",
        xaxis_title="Date", yaxis_title="Contribution / GA (%)",
        template="plotly_white",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.25,
                    xanchor="center", x=0.5),
    )
    st.plotly_chart(fig_contrib, use_container_width=True)

with tab4:
    # Tableau de taux de croissance (GA) trimestriels sur les 8 derniers trimestres
    dates_idx = pd.to_datetime(result_df.index)
    q_cemac = quarterly_cemac(result_df, dates_idx)
    st.session_state["cemac_quarterly"] = q_cemac

    # Recalculer proprement les GA trimestriels par pays
    ga_trim_data = {"Trimestre": q_cemac["Trimestre"].astype(str).tolist()}
    for code in COUNTRY_CODES:
        if code in q_cemac.columns:
            ga_pays_t = (q_cemac[code] / q_cemac[code].shift(4) - 1) * 100
            ga_trim_data[COUNTRY_NAMES.get(code, code)] = ga_pays_t.round(2).tolist()
    ga_trim_data["CEMAC"] = q_cemac["GA_Trim"].round(2).tolist() if "GA_Trim" in q_cemac.columns else []

    ga_trim_df = pd.DataFrame(ga_trim_data)
    # Afficher les 8 derniers trimestres
    n_show = min(8, len(ga_trim_df))
    st.subheader("Taux de croissance trimestriels (GA, %)")
    st.dataframe(ga_trim_df.tail(n_show), use_container_width=True, hide_index=True)

with tab5:
    st.dataframe(result_df, use_container_width=True)

# Trimestriel complet
st.subheader("Résultats trimestriels")
if "cemac_quarterly" not in st.session_state:
    dates_idx = pd.to_datetime(result_df.index)
    q_cemac = quarterly_cemac(result_df, dates_idx)
    st.session_state["cemac_quarterly"] = q_cemac
q_cemac = st.session_state["cemac_quarterly"]
st.dataframe(q_cemac, use_container_width=True, hide_index=True)

# ── Export ────────────────────────────────────────────────────────────────
st.header("6. Export")
if st.button("📥 Exporter CEMAC", key="export_cemac"):
    data = write_cemac_excel(
        result_df, q_cemac,
        st.session_state.get("cemac_computed_poids", POIDS_PIB),
        template_path=CEMAC_TEMPLATE,
    )
    download_button(
        data,
        "ICAE_CEMAC_Consolide_OUT.xlsx",
        "📥 Télécharger le fichier CEMAC",
    )
