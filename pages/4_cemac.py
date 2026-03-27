"""Module 4 — Agrégation CEMAC (ICAE + PIB Nowcast)."""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from pathlib import Path

from config import (CONSOLIDES, COUNTRY_CODES, COUNTRY_NAMES,
                    PIB_2014, PIB_REF_YEAR, POIDS_PIB, CLOUD_MODE, CEMAC_TEMPLATE)
from core.cemac_engine import compute_icae_cemac, quarterly_cemac
from core.icae_engine import run_icae_pipeline
from io_utils.excel_reader import (
    list_sheets, load_country_file, read_consignes, read_codification,
    read_donnees_calcul, rename_columns_to_codes,
)
from io_utils.excel_writer import write_cemac_excel
from ui.charts import chart_icae_monthly, chart_ga_bars
from ui.components import download_button

_COUNTRY_COLORS = {
    "CMR": "#4472C4", "RCA": "#ED7D31", "CNG": "#A5A5A5",
    "GAB": "#FFC000", "GNQ": "#5B9BD5", "TCD": "#70AD47",
}


def _process_country_file(source):
    """Traite un fichier ICAE pays (chemin ou file-like) → (icae_series, dates, err) ou None."""
    try:
        if hasattr(source, 'seek'):
            source.seek(0)
        _sh = list_sheets(source)
        if hasattr(source, 'seek'):
            source.seek(0)
        if "Donnees_calcul" not in _sh:
            return None, None, f"Feuille 'Donnees_calcul' absente. Feuilles trouvées : {_sh}"
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
        return icae, dates, None
    except Exception as exc:
        import traceback
        return None, None, traceback.format_exc()


st.title("🌍 Module 4 — Agrégation CEMAC")

# ── Source ────────────────────────────────────────────────────────────────
st.header("1. Données ICAE")

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
        st.info("Uploadez au moins un fichier consolidé ICAE pour continuer.")
        # Ne pas appeler st.stop() ici — laisser Streamlit afficher les uploaders

    st.caption(f"📁 {len(uploaded_files)}/{len(COUNTRY_CODES)} fichiers uploadés")
    progress = st.progress(0)
    for i, (code, f) in enumerate(uploaded_files.items()):
        icae_s, dates_s, err = _process_country_file(f)
        if icae_s is not None:
            icae_dict[code], dates_dict[code] = icae_s, dates_s
        else:
            st.warning(f"⚠️ Erreur pour {code}")
            if err:
                with st.expander(f"Détails erreur {code}", expanded=False):
                    st.code(err)
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
        icae_s, dates_s, err = _process_country_file(fpath)
        if icae_s is not None:
            icae_dict[code], dates_dict[code] = icae_s, dates_s
        else:
            st.warning(f"⚠️ Erreur pour {code}")
            if err:
                with st.expander(f"Détails erreur {code}", expanded=False):
                    st.code(err)
        progress.progress((i + 1) / len(COUNTRY_CODES))

    if icae_dict:
        st.session_state["icae_monthly"] = icae_dict
        st.success(f"✅ ICAE calculés pour : {', '.join(icae_dict.keys())}")

if not icae_dict:
    st.warning("Aucun ICAE disponible.")
    st.stop()

# ── Contrôle qualité : valeurs hors plage [50–200] ───────────────────────
_out_of_range = []
for _code, _ser in icae_dict.items():
    try:
        _vals = pd.to_numeric(_ser, errors="coerce").dropna()
        if len(_vals) > 0 and (_vals.max() > 200 or _vals.min() < 50):
            _out_of_range.append(
                f"**{COUNTRY_NAMES.get(_code, _code)}** ({_code}) : "
                f"min={_vals.min():.1f}, max={_vals.max():.1f}"
            )
    except Exception:
        pass

if _out_of_range:
    with st.expander("⚠️ Valeurs ICAE hors plage [50–200] détectées — cliquer pour détails",
                     expanded=True):
        st.warning(
            "Un ICAE correctement calculé (base 100) reste normalement entre **50 et 200**. "
            "Les séries suivantes présentent des valeurs atypiques :"
        )
        for _msg in _out_of_range:
            st.markdown(f"- {_msg}")
        st.caption(
            "Causes possibles : année de base incorrecte dans le fichier consolidé, "
            "variable aberrante non neutralisée, ou fichier non mis à jour."
        )

# ── Période commune ──────────────────────────────────────────────────────
st.header("2. Période de calcul")

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

# Option : recalculer les poids depuis un fichier PIB externe
with st.expander("📂 Recalculer les poids depuis un fichier PIB", expanded=False):
    st.caption(
        "Uploadez un fichier Excel contenant une colonne **Pays/Code** et des colonnes "
        "d'années (ex : 2014, 2015, …). Choisissez l'année de référence pour recalculer les poids."
    )
    _pib_file = st.file_uploader(
        "Fichier PIB (xlsx)", type=["xlsx"], key="cemac_pib_upload"
    )
    if _pib_file is not None:
        try:
            _pib_sheets = list_sheets(_pib_file)
            _pib_file.seek(0)
            _pib_sheet = st.selectbox("Feuille", _pib_sheets, key="cemac_pib_sheet")
            _pib_file.seek(0)
            _pib_raw = pd.read_excel(_pib_file, sheet_name=_pib_sheet, engine="openpyxl")
            st.dataframe(_pib_raw.head(8), use_container_width=True)

            # Détecter la colonne pays (premier mot-clé)
            _id_col = _pib_raw.columns[0]
            _year_cols = [c for c in _pib_raw.columns[1:]
                          if str(c).strip().isdigit() or
                          (isinstance(c, (int, float)) and not np.isnan(c))]
            if _year_cols:
                _ref_year = st.selectbox(
                    "Année de référence pour les poids",
                    [str(int(float(c))) for c in _year_cols],
                    key="cemac_pib_year",
                )
                if st.button("✅ Appliquer les poids", key="apply_pib_weights"):
                    _chosen_col = None
                    for c in _pib_raw.columns:
                        if str(c).strip() == _ref_year or str(int(float(c))) == _ref_year:
                            _chosen_col = c
                            break
                    if _chosen_col is not None:
                        _new_pibs = {}
                        for _, row_p in _pib_raw.iterrows():
                            code_raw = str(row_p[_id_col]).strip().upper()
                            # Match souple : code pays ou nom pays
                            matched = None
                            for cc in COUNTRY_CODES:
                                if cc in code_raw or COUNTRY_NAMES[cc].upper() in code_raw:
                                    matched = cc
                                    break
                            if matched:
                                try:
                                    _new_pibs[matched] = float(row_p[_chosen_col])
                                except (ValueError, TypeError):
                                    pass
                        if _new_pibs:
                            st.session_state["cemac_custom_pib"] = _new_pibs
                            st.session_state["cemac_pib_year"] = _ref_year
                            st.success(f"✅ PIB {_ref_year} appliqués pour : {list(_new_pibs.keys())}")
        except Exception as _e:
            st.warning(f"Erreur lecture fichier PIB : {_e}")

# Initialiser les valeurs PIB (custom ou défaut)
_base_pib = st.session_state.get("cemac_custom_pib", PIB_2014)
_pib_ref_year = st.session_state.get("cemac_pib_year", PIB_REF_YEAR)
_pib_label = f"PIB {_pib_ref_year} (Mds FCFA)"

poids_data = []
for code in COUNTRY_CODES:
    poids_data.append({
        "Code": code,
        "Pays": COUNTRY_NAMES[code],
        _pib_label: _base_pib.get(code, PIB_2014[code]),
        "Poids (%)": round(_base_pib.get(code, PIB_2014[code]) /
                           sum(_base_pib.get(c, PIB_2014[c]) for c in COUNTRY_CODES) * 100, 2),
    })

poids_df = pd.DataFrame(poids_data)
edited_poids = st.data_editor(poids_df, key="cemac_poids",
                              use_container_width=True, num_rows="fixed")

pib_edited = dict(zip(edited_poids["Code"], edited_poids[_pib_label]))
total_pib = sum(pib_edited.values())
poids = {code: pib_edited[code] / total_pib for code in pib_edited}

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

st.header("5. Résultats ICAE CEMAC")

st.subheader("📅 Fenêtre d'affichage")
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

_poids_used = st.session_state.get("cemac_computed_poids", POIDS_PIB)

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📈 ICAE mensuel",
    "📊 GA mensuel",
    "📊 Contributions pays (mensuel)",
    "📈 ICAE trimestriel",
    "📊 GA trimestriel",
    "📊 Contributions pays (trimestriel)",
    "📋 Données",
])

with tab1:
    fig = chart_icae_monthly(
        _disp_df.index, _disp_df["ICAE_CEMAC"],
        title="ICAE CEMAC — Agrégé pondéré",
    )
    for code in COUNTRY_CODES:
        if code in _disp_df.columns and code in pays_inclus:
            fig.add_scatter(x=_disp_df.index, y=_disp_df[code],
                           mode="lines", name=COUNTRY_NAMES.get(code, code),
                           opacity=0.5)
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    fig_ga = chart_ga_bars(
        _disp_df.index, _disp_df["GA"],
        title="ICAE CEMAC — Glissement annuel mensuel (%)",
    )
    st.plotly_chart(fig_ga, use_container_width=True)

with tab3:
    # Calculer les contributions sur le result_df COMPLET pour éviter les NaN
    # dus au shift(12) sur les bords de la fenêtre filtrée
    _full_dates = pd.to_datetime(result_df.index)
    fig_contrib = go.Figure()
    for code in COUNTRY_CODES:
        if code in result_df.columns and code in pays_inclus:
            ga_full = (result_df[code] / result_df[code].shift(12) - 1) * 100
            ga_full.index = _full_dates
            contrib_full = ga_full * _poids_used.get(code, 0)
            # Filtrer sur la fenêtre d'affichage
            contrib_disp = contrib_full[_graph_mask.values]
            fig_contrib.add_trace(go.Bar(
                x=_disp_df.index, y=contrib_disp,
                name=f"{COUNTRY_NAMES.get(code, code)} ({round(_poids_used.get(code, 0) * 100, 1)}%)",
                marker_color=_COUNTRY_COLORS.get(code, "#888"),
            ))
    # Courbe GA CEMAC
    fig_contrib.add_trace(go.Scatter(
        x=_disp_df.index, y=_disp_df["GA"],
        mode="lines+markers", name="GA CEMAC (%)",
        line=dict(color="#C00000", width=3),
    ))
    fig_contrib.update_layout(
        barmode="relative",
        title="Contributions pays au GA de l'ICAE CEMAC — mensuel (%)",
        xaxis_title="Date", yaxis_title="Contribution / GA (%)",
        template="plotly_white", hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.28,
                    xanchor="center", x=0.5),
    )
    st.plotly_chart(fig_contrib, use_container_width=True)

# ── Trimestriel ──────────────────────────────────────────────────────────
dates_idx_full = pd.to_datetime(result_df.index)
q_cemac = quarterly_cemac(result_df, dates_idx_full)
st.session_state["cemac_quarterly"] = q_cemac

# Filtrer les trimestres sur la fenêtre d'affichage
_q_dates = pd.to_datetime([
    str(t) for t in q_cemac.get("Trimestre", q_cemac.index)
], errors="coerce")
# Si Trimestre est une colonne text, parser proprement
if "Trimestre" in q_cemac.columns:
    _q_start_mask = pd.Series(True, index=q_cemac.index)
    _q_end_mask = pd.Series(True, index=q_cemac.index)
    _q_disp = q_cemac.copy()
else:
    _q_disp = q_cemac.copy()

with tab4:
    # ICAE trimestriel
    _icae_col = "ICAE_Trim" if "ICAE_Trim" in q_cemac.columns else (
        "icae_cemac_trim" if "icae_cemac_trim" in q_cemac.columns else
        q_cemac.select_dtypes("number").columns[0] if not q_cemac.select_dtypes("number").empty
        else None
    )
    _trim_col = "Trimestre" if "Trimestre" in q_cemac.columns else q_cemac.columns[0]
    if _icae_col:
        fig_q = go.Figure()
        fig_q.add_trace(go.Scatter(
            x=_q_disp[_trim_col].astype(str), y=_q_disp[_icae_col],
            mode="lines+markers", name="ICAE CEMAC trim.",
            line=dict(color="#1F4E79", width=2),
        ))
        fig_q.update_layout(
            title="ICAE CEMAC — Trimestriel",
            xaxis_title="Trimestre", yaxis_title="Indice",
            template="plotly_white",
        )
        st.plotly_chart(fig_q, use_container_width=True)
    st.dataframe(_q_disp, use_container_width=True, hide_index=True)

with tab5:
    # GA trimestriel CEMAC + pays
    _ga_trim_col = "GA_Trim" if "GA_Trim" in q_cemac.columns else None
    fig_ga_q = go.Figure()
    if _ga_trim_col:
        fig_ga_q.add_trace(go.Bar(
            x=_q_disp[_trim_col].astype(str), y=_q_disp[_ga_trim_col],
            name="GA CEMAC trim. (%)",
            marker_color="#C00000",
        ))
    # GA trimestriel par pays
    for code in COUNTRY_CODES:
        if code in q_cemac.columns and code in pays_inclus:
            _ga_p = (q_cemac[code] / q_cemac[code].shift(4) - 1) * 100
            fig_ga_q.add_trace(go.Scatter(
                x=_q_disp[_trim_col].astype(str), y=_ga_p.values,
                mode="lines+markers",
                name=f"{COUNTRY_NAMES.get(code, code)}",
                marker_color=_COUNTRY_COLORS.get(code, "#888"),
                opacity=0.75,
            ))
    fig_ga_q.update_layout(
        barmode="overlay",
        title="GA de l'ICAE CEMAC — Trimestriel (%)",
        xaxis_title="Trimestre", yaxis_title="GA (%)",
        template="plotly_white", hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.28,
                    xanchor="center", x=0.5),
    )
    st.plotly_chart(fig_ga_q, use_container_width=True)

    # Tableau GA trimestriel
    ga_trim_data = {_trim_col: _q_disp[_trim_col].astype(str).tolist()}
    for code in COUNTRY_CODES:
        if code in q_cemac.columns:
            ga_trim_data[COUNTRY_NAMES.get(code, code)] = (
                (q_cemac[code] / q_cemac[code].shift(4) - 1) * 100
            ).round(2).tolist()
    if _ga_trim_col:
        ga_trim_data["CEMAC"] = q_cemac[_ga_trim_col].round(2).tolist()
    st.dataframe(pd.DataFrame(ga_trim_data).tail(8),
                 use_container_width=True, hide_index=True)

with tab6:
    # Contributions pays trimestrielles
    fig_contrib_q = go.Figure()
    for code in COUNTRY_CODES:
        if code in q_cemac.columns and code in pays_inclus:
            ga_p_q = (q_cemac[code] / q_cemac[code].shift(4) - 1) * 100
            contrib_p_q = ga_p_q * _poids_used.get(code, 0)
            fig_contrib_q.add_trace(go.Bar(
                x=_q_disp[_trim_col].astype(str), y=contrib_p_q.values,
                name=f"{COUNTRY_NAMES.get(code, code)} ({round(_poids_used.get(code, 0) * 100, 1)}%)",
                marker_color=_COUNTRY_COLORS.get(code, "#888"),
            ))
    if _ga_trim_col:
        fig_contrib_q.add_trace(go.Scatter(
            x=_q_disp[_trim_col].astype(str), y=_q_disp[_ga_trim_col],
            mode="lines+markers", name="GA CEMAC trim. (%)",
            line=dict(color="#C00000", width=3),
        ))
    fig_contrib_q.update_layout(
        barmode="relative",
        title="Contributions pays au GA de l'ICAE CEMAC — trimestriel (%)",
        xaxis_title="Trimestre", yaxis_title="Contribution / GA (%)",
        template="plotly_white", hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.28,
                    xanchor="center", x=0.5),
    )
    st.plotly_chart(fig_contrib_q, use_container_width=True)

with tab7:
    st.dataframe(result_df, use_container_width=True)

# ── PIB Nowcast CEMAC ────────────────────────────────────────────────────
_has_nowcast = (
    "nowcast_results" in st.session_state
    and "nowcast_pib" in st.session_state
    and len(st.session_state["nowcast_results"]) > 0
)

if _has_nowcast:
    st.header("6. PIB Nowcast CEMAC")
    st.info(
        "Les données Nowcast suivantes proviennent du Module 3. "
        "Seuls les pays ayant un Nowcast calculé sont affichés."
    )

    _nw_results = st.session_state["nowcast_results"]   # {pays: {modele: {forecast, ...}}}
    _nw_pib = st.session_state["nowcast_pib"]           # {pays: pd.Series PIB obs}

    # Sélection du modèle à afficher
    _all_models = sorted({m for res in _nw_results.values() for m in res.keys()})
    _sel_model = st.selectbox("Modèle Nowcast à afficher",
                              _all_models, key="cemac_nw_model") if _all_models else None

    # Onglets mensuel/trimestriel pour le Nowcast
    _nwtab1, _nwtab2 = st.tabs(["📈 PIB Nowcast — Niveaux", "📊 PIB Nowcast — Taux de croissance (GA)"])

    with _nwtab1:
        fig_nw = go.Figure()
        for code, pib_obs in _nw_pib.items():
            if code not in pays_inclus:
                continue
            fig_nw.add_trace(go.Scatter(
                x=[str(p) for p in pib_obs.index],
                y=pib_obs.values,
                mode="lines", name=f"PIB obs. {COUNTRY_NAMES.get(code, code)}",
                line=dict(color=_COUNTRY_COLORS.get(code, "#888"), dash="solid"),
            ))
            if _sel_model and code in _nw_results and _sel_model in _nw_results[code]:
                fc = _nw_results[code][_sel_model]["forecast"]
                fig_nw.add_trace(go.Scatter(
                    x=[str(p) for p in fc.index],
                    y=fc.values,
                    mode="lines+markers",
                    name=f"Nowcast {code} ({_sel_model})",
                    line=dict(color=_COUNTRY_COLORS.get(code, "#888"), dash="dot"),
                ))
        fig_nw.update_layout(
            title="PIB observé et Nowcast par pays (niveaux)",
            xaxis_title="Trimestre", yaxis_title="PIB",
            template="plotly_white", hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=-0.3,
                        xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_nw, use_container_width=True)

    with _nwtab2:
        fig_ga_nw = go.Figure()
        for code, pib_obs in _nw_pib.items():
            if code not in pays_inclus:
                continue
            ga_obs = (pib_obs / pib_obs.shift(4) - 1) * 100
            fig_ga_nw.add_trace(go.Scatter(
                x=[str(p) for p in ga_obs.index],
                y=ga_obs.values,
                mode="lines+markers",
                name=f"GA PIB obs. {COUNTRY_NAMES.get(code, code)}",
                line=dict(color=_COUNTRY_COLORS.get(code, "#888"), dash="solid"),
            ))
            if _sel_model and code in _nw_results and _sel_model in _nw_results[code]:
                fc = _nw_results[code][_sel_model]["forecast"]
                pib_all = pd.concat([pib_obs, fc[fc.index > pib_obs.index[-1]]])
                ga_fc = (pib_all / pib_all.shift(4) - 1) * 100
                ga_fc_only = ga_fc[fc.index]
                fig_ga_nw.add_trace(go.Scatter(
                    x=[str(p) for p in ga_fc_only.index],
                    y=ga_fc_only.values,
                    mode="lines+markers",
                    name=f"GA Nowcast {code} ({_sel_model})",
                    line=dict(color=_COUNTRY_COLORS.get(code, "#888"), dash="dot"),
                ))
        fig_ga_nw.update_layout(
            title="Taux de croissance annuel (GA) — PIB observé et Nowcast",
            xaxis_title="Trimestre", yaxis_title="GA (%)",
            template="plotly_white", hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=-0.3,
                        xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_ga_nw, use_container_width=True)
else:
    st.info(
        "💡 Lancez le **Module 3 — Nowcast** pour afficher ici les PIB Nowcast "
        "par pays et les agréger au niveau CEMAC."
    )

# ── Export ────────────────────────────────────────────────────────────────
st.header("7. Export")
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
