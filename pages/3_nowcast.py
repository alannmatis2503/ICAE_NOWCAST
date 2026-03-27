"""Module 3 — Nowcast."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import COUNTRY_CODES, COUNTRY_NAMES, NOWCAST_MODELS, get_default_agg_type
from core.nowcast_engine import run_nowcast
from core.quarterly import agg_m_to_q
from core.tempdisagg import disaggregate_annual_to_quarterly
from io_utils.excel_reader import list_sheets, read_codification
from io_utils.excel_writer import write_nowcast_excel
from ui.charts import chart_nowcast, chart_ga_nowcast
from ui.components import download_button

st.title("🔮 Module 3 — Nowcast")


# ── Helpers ───────────────────────────────────────────────────────────────
def _get_codif_vars(codif_df, available_cols):
    """Retourne la liste des variables actives depuis la feuille Codification.

    Correspond par Code ET par Label (insensible à la casse et aux espaces
    superflus) pour gérer les classeurs dont les colonnes sont des libellés.
    """
    if codif_df is None or "Code" not in codif_df.columns:
        return []
    active_codes: set = set()
    for _, row in codif_df.iterrows():
        try:
            if float(row.get("PRIOR", 0)) > 0:
                statut = str(row.get("Statut", "Actif")).lower().strip()
                if not statut.startswith("inact"):
                    active_codes.add(str(row["Code"]).strip().lower())
                    if "Label" in codif_df.columns and pd.notna(row.get("Label")):
                        active_codes.add(str(row["Label"]).strip().lower())
        except (ValueError, TypeError):
            pass
    return [c for c in available_cols if str(c).strip().lower() in active_codes]


def _annualize_monthly(df, date_col, value_cols):
    """Annualise des données mensuelles : sum (flux), mean (taux), last (stock).

    Retourne un DataFrame indexé par année.
    """
    tmp = df[[date_col] + value_cols].copy()
    tmp[date_col] = pd.to_datetime(tmp[date_col], errors="coerce")
    tmp = tmp.dropna(subset=[date_col])
    tmp["_year"] = tmp[date_col].dt.year
    num = tmp[value_cols].apply(pd.to_numeric, errors="coerce")
    num["_year"] = tmp["_year"]

    result_parts = []
    for v in value_cols:
        agg_type = get_default_agg_type(v)
        if agg_type == "last":
            agg_fn = "last"
        elif agg_type == "mean":
            agg_fn = "mean"
        else:
            agg_fn = "sum"
        s = num.groupby("_year")[v].agg(agg_fn)
        result_parts.append(s)

    return pd.concat(result_parts, axis=1)


# ── Étape 0 : Sélection des variables ────────────────────────────────────
with st.expander("⚙️ Étape 0 — Sélection des variables candidates", expanded=False):
    st.markdown("""
    **Phase de sélection structurelle** (à réaliser tous les 2-3 ans) :
    - **Codification** : utilise les variables candidates de la feuille Codification du classeur (défaut)
    - **ACP + Corrélation** : combine l'ACP (avec le PIB comme variable supplémentaire) et la corrélation
      avec le PIB pour identifier les meilleures variables. Les données mensuelles sont annualisées
      automatiquement (somme pour les flux, moyenne les taux, fin de période pour les stocks).
    """)

    var_sel_method = st.radio(
        "Méthode de sélection",
        ["Depuis la feuille Codification",
         "ACP + Corrélation avec le PIB",
         "Aucune (utiliser toutes les variables)"],
        key="now_var_sel_method",
    )
    st.session_state["_now_var_sel"] = var_sel_method

# ── Données HF ───────────────────────────────────────────────────────────
st.header("1. Indicateurs haute fréquence")

# Auto-sélection si "séries prolongées" depuis Module 2
_use_extended = st.session_state.pop("_nowcast_use_extended", False)
if _use_extended:
    # Purger les anciens résultats Nowcast pour forcer un recalcul avec
    # les séries prolongées — évite de télécharger des résultats périmés
    for _k in ("nowcast_results", "nowcast_pib"):
        st.session_state.pop(_k, None)
_hf_sources = ["Upload d'un fichier", "Données du Module 1"]
_default_src = 1 if (_use_extended and "donnees_calcul" in st.session_state) else 0

hf_source = st.radio(
    "Source des indicateurs HF",
    _hf_sources,
    index=_default_src,
    horizontal=True,
    key="now_hf_source",
)

hf_df = None
_now_codif = None

if hf_source == "Données du Module 1" and "donnees_calcul" in st.session_state:
    available = list(st.session_state["donnees_calcul"].keys())
    country = st.selectbox("Pays", available, key="now_country")
    hf_df = st.session_state["donnees_calcul"][country]
    _ext_label = " (séries prolongées)" if _use_extended else ""
    st.success(f"✅ Données HF chargées depuis Module 1 ({country}){_ext_label}")
    _codif_dict = st.session_state.get("codification", {})
    if country in _codif_dict:
        _now_codif = _codif_dict[country]
else:
    uploaded_hf = st.file_uploader("Fichier des indicateurs mensuels", type=["xlsx"],
                                    key="now_hf_upload")
    if uploaded_hf:
        sheets = list_sheets(uploaded_hf)
        uploaded_hf.seek(0)
        sheet_hf = st.selectbox("Feuille HF", sheets, key="now_hf_sheet")
        uploaded_hf.seek(0)
        hf_df = pd.read_excel(uploaded_hf, sheet_name=sheet_hf, engine="openpyxl")
        # Lire la Codification si disponible dans le même classeur
        uploaded_hf.seek(0)
        if "Codification" in sheets:
            try:
                _now_codif = read_codification(uploaded_hf)
                uploaded_hf.seek(0)
            except Exception:
                pass
        country = "XXX"
        for c in COUNTRY_CODES:
            if c in uploaded_hf.name:
                country = c
                break

if hf_df is None:
    if ("nowcast_results" in st.session_state
            and "nowcast_pib" in st.session_state):
        # Afficher les résultats précédents sans re-charger les données
        _c_prev = list(st.session_state["nowcast_results"].keys())[0]
        _res_prev = st.session_state["nowcast_results"][_c_prev]
        _pib_prev = st.session_state["nowcast_pib"][_c_prev]
        st.info(
            f"⚠️ Résultats du dernier Nowcast disponibles pour **{_c_prev}**. "
            "Rechargez les indicateurs HF pour lancer une nouvelle analyse."
        )
        st.header(f"Résultats enregistrés — {_c_prev}")
        _tab1, _tab2, _tab3 = st.tabs(
            ["📈 Graphiques", "📊 Performance", "📉 GA Nowcast"]
        )
        with _tab1:
            from ui.charts import chart_nowcast
            st.plotly_chart(
                chart_nowcast(_pib_prev, _res_prev,
                              title=f"PIB observé vs Nowcasts — {_c_prev}"),
                use_container_width=True,
            )
        with _tab2:
            _perf = []
            for _nm, _r in _res_prev.items():
                _m = _r["metrics"]
                _perf.append({
                    "Modèle": _nm,
                    "RMSE (in)": round(_m["in_sample"].get("rmse", np.nan), 2),
                    "MAE (in)": round(_m["in_sample"].get("mae", np.nan), 2),
                    "MAPE (in)": round(_m["in_sample"].get("mape", np.nan), 2),
                    "RMSE (out)": round(_m["out_sample"].get("rmse", np.nan), 2),
                    "MAE (out)": round(_m["out_sample"].get("mae", np.nan), 2),
                    "MAPE (out)": round(_m["out_sample"].get("mape", np.nan), 2),
                    "Corrélation": round(_r.get("correlation", np.nan), 4),
                })
            st.dataframe(pd.DataFrame(_perf), use_container_width=True,
                         hide_index=True)
        with _tab3:
            from ui.charts import chart_ga_nowcast
            st.plotly_chart(
                chart_ga_nowcast(_pib_prev, _res_prev,
                                 title=f"GA du PIB et des Nowcasts — {_c_prev}"),
                use_container_width=True,
            )
    else:
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

    # Déterminer les variables par défaut selon la méthode de sélection
    _var_sel = st.session_state.get("_now_var_sel", "Depuis la feuille Codification")

    # Forcer le reset du multiselect quand la méthode OU la source change
    _current_ctx = f"{_var_sel}|{hf_source}|{country}"
    _prev_ctx = st.session_state.get("_now_var_ctx_prev")
    if _prev_ctx is not None and _prev_ctx != _current_ctx:
        st.session_state.pop("now_hf_vars", None)
    st.session_state["_now_var_ctx_prev"] = _current_ctx

    if _var_sel == "Depuis la feuille Codification":
        codif_filtered = _get_codif_vars(_now_codif, non_date)
        if codif_filtered:
            default_hf = codif_filtered
            st.info(f"🎯 {len(codif_filtered)} variables candidates depuis la Codification")
        else:
            default_hf = non_date
            if _now_codif is not None:
                st.warning(
                    "⚠️ Aucune correspondance entre la Codification et les colonnes "
                    "du fichier. Vérifiez que les noms de colonnes correspondent "
                    "aux codes ou labels de la Codification. Toutes les variables "
                    "sont affichées par défaut."
                )
            else:
                st.warning(
                    "⚠️ Aucune feuille Codification trouvée. "
                    "Toutes les variables sont affichées par défaut."
                )
    elif _var_sel == "ACP + Corrélation avec le PIB":
        default_hf = non_date
    else:
        default_hf = non_date

    hf_vars = st.multiselect("Variables HF", non_date, default=default_hf,
                             key="now_hf_vars")

if not hf_vars:
    st.warning("Sélectionnez au moins une variable HF.")
    st.stop()

# ── Analyse ACP + Corrélation combinée ───────────────────────────────────
_var_sel = st.session_state.get("_now_var_sel", "Depuis la feuille Codification")

if _var_sel == "ACP + Corrélation avec le PIB":
    with st.expander("📊 Résultats ACP + Corrélation avec le PIB", expanded=True):
        import plotly.express as px
        import plotly.graph_objects as go
        from sklearn.preprocessing import StandardScaler
        from sklearn.decomposition import PCA

        st.markdown("""
        Les données mensuelles sont **annualisées** automatiquement avant l'ACP :
        - **Flux** (production, exports…) → somme annuelle
        - **Taux** (IPC, ratio…) → moyenne annuelle
        - **Stocks** (masse monétaire…) → valeur de fin de période
        Le **PIB** est traité comme **variable supplémentaire** (projeté sur les axes ACP sans influencer le calcul).
        """)

        # Annualiser les données mensuelles pour l'ACP
        annual_df = _annualize_monthly(hf_df, date_col, hf_vars)
        annual_df = annual_df.apply(pd.to_numeric, errors="coerce").dropna()

        # Charger le PIB si disponible en session
        pib_for_acp = None
        st.markdown("**PIB annuel** (variable supplémentaire pour l'ACP) :")
        pib_acp_upload = st.file_uploader("Fichier PIB annuel (optionnel)",
                                          type=["xlsx"], key="pib_acp_upload")
        if pib_acp_upload:
            _sheets_pib = list_sheets(pib_acp_upload)
            pib_acp_upload.seek(0)
            _sheet = st.selectbox("Feuille", _sheets_pib, key="pib_acp_sheet")
            pib_acp_upload.seek(0)
            _raw = pd.read_excel(pib_acp_upload, sheet_name=_sheet, engine="openpyxl")
            _cols = list(_raw.columns)
            _c1, _c2 = st.columns(2)
            with _c1:
                _d = st.selectbox("Colonne année", _cols, key="pib_acp_dcol")
            with _c2:
                _v = st.selectbox("Colonne PIB", [c for c in _cols if c != _d],
                                  key="pib_acp_vcol")
            _ps = _raw[[_d, _v]].dropna()
            pib_for_acp = pd.Series(_ps[_v].values,
                                    index=_ps[_d].astype(int).values,
                                    name="PIB")

        if len(annual_df) > 2 and len(hf_vars) > 1:
            scaler = StandardScaler()
            X_scaled = scaler.fit_transform(annual_df.values)
            n_comp = min(len(hf_vars), 5, len(annual_df) - 1)
            pca = PCA(n_components=n_comp)
            pcs = pca.fit_transform(X_scaled)

            # ── Variance expliquée ──
            st.subheader("Variance expliquée par composante")
            var_df = pd.DataFrame({
                "Composante": [f"PC{i+1}" for i in range(n_comp)],
                "Variance expliquée (%)": (pca.explained_variance_ratio_ * 100).round(2),
                "Cumulée (%)": (pca.explained_variance_ratio_.cumsum() * 100).round(2),
            })
            st.dataframe(var_df, use_container_width=True, hide_index=True)

            # ── Loadings ──
            loadings = pd.DataFrame(
                pca.components_.T,
                index=hf_vars,
                columns=[f"PC{i+1}" for i in range(n_comp)],
            ).round(4)
            st.subheader("Loadings (contributions aux composantes)")
            st.dataframe(loadings, use_container_width=True)

            # ── Cercle de corrélation avec PIB supplémentaire ──
            st.subheader("Cercle de corrélation (PC1 × PC2)")
            sqrt_ev = np.sqrt(pca.explained_variance_[:2])
            corr_circle = loadings[["PC1", "PC2"]].copy()
            corr_circle["PC1"] *= sqrt_ev[0]
            corr_circle["PC2"] *= sqrt_ev[1]

            fig_circle = go.Figure()
            # Cercle unité
            theta = np.linspace(0, 2 * np.pi, 100)
            fig_circle.add_trace(go.Scatter(
                x=np.cos(theta), y=np.sin(theta),
                mode="lines", line=dict(color="grey", dash="dash"),
                showlegend=False,
            ))
            # Variables actives
            for var_name in corr_circle.index:
                x_v, y_v = corr_circle.loc[var_name, "PC1"], corr_circle.loc[var_name, "PC2"]
                fig_circle.add_trace(go.Scatter(
                    x=[0, x_v], y=[0, y_v], mode="lines+text",
                    text=["", var_name], textposition="top center",
                    line=dict(color="#1F4E79"), showlegend=False,
                ))
            # PIB en variable supplémentaire
            if pib_for_acp is not None:
                common_years = annual_df.index.intersection(pib_for_acp.index)
                if len(common_years) >= 3:
                    pib_vals = pib_for_acp.loc[common_years].values.astype(float)
                    pib_scaled = (pib_vals - pib_vals.mean()) / (pib_vals.std() + 1e-9)
                    # Projeter PIB sur les axes PCA
                    X_common = annual_df.loc[common_years].values
                    X_common_scaled = scaler.transform(X_common)
                    pib_corr_pc1 = np.corrcoef(pib_scaled, pcs[:len(common_years), 0][:len(pib_scaled)])[0, 1]
                    pib_corr_pc2 = np.corrcoef(pib_scaled, pcs[:len(common_years), 1][:len(pib_scaled)])[0, 1]
                    fig_circle.add_trace(go.Scatter(
                        x=[0, pib_corr_pc1], y=[0, pib_corr_pc2],
                        mode="lines+text", text=["", "★ PIB"],
                        textposition="top center",
                        line=dict(color="#E26B0A", width=3),
                        showlegend=False,
                    ))
                    fig_circle.add_trace(go.Scatter(
                        x=[pib_corr_pc1], y=[pib_corr_pc2],
                        mode="markers", marker=dict(color="#E26B0A", size=12, symbol="star"),
                        name="PIB (suppl.)",
                    ))

            fig_circle.update_layout(
                xaxis_title=f"PC1 ({pca.explained_variance_ratio_[0]*100:.1f}%)",
                yaxis_title=f"PC2 ({pca.explained_variance_ratio_[1]*100:.1f}%)",
                xaxis=dict(range=[-1.1, 1.1], scaleanchor="y"),
                yaxis=dict(range=[-1.1, 1.1]),
                height=600,
            )
            st.plotly_chart(fig_circle, use_container_width=True)

            # ── Corrélation des variables avec le PIB (si disponible) ──
            if pib_for_acp is not None and len(common_years) >= 3:
                st.subheader("Corrélation des variables annualisées avec le PIB")
                corr_pib = {}
                for v in hf_vars:
                    if v in annual_df.columns:
                        vals = annual_df.loc[common_years, v].values.astype(float)
                        r = np.corrcoef(pib_vals, vals)[0, 1]
                        if not np.isnan(r):
                            corr_pib[v] = round(r, 4)
                corr_pib_sorted = dict(sorted(corr_pib.items(),
                                              key=lambda x: abs(x[1]), reverse=True))
                corr_df = pd.DataFrame({
                    "Variable": corr_pib_sorted.keys(),
                    "Corrélation avec PIB": corr_pib_sorted.values(),
                    "|r|": [abs(v) for v in corr_pib_sorted.values()],
                })
                st.dataframe(corr_df, use_container_width=True, hide_index=True)

                # Heatmap corrélation
                fig_heat = px.bar(
                    corr_df, x="Variable", y="Corrélation avec PIB",
                    color="|r|", color_continuous_scale="RdBu_r",
                    title="Corrélation de chaque variable avec le PIB",
                )
                fig_heat.add_hline(y=0, line_dash="dash")
                st.plotly_chart(fig_heat, use_container_width=True)

                # Proposer les variables les plus corrélées
                threshold_corr = st.slider("Seuil |corrélation| pour sélection",
                                           0.0, 1.0, 0.5, step=0.05,
                                           key="corr_pib_thresh")
                selected_from_corr = [v for v, r in corr_pib.items()
                                      if abs(r) >= threshold_corr]
                st.info(f"**{len(selected_from_corr)}** variables avec |r| ≥ {threshold_corr}")

            # Top contributeurs PC1
            abs_load = loadings["PC1"].abs().sort_values(ascending=False)
            threshold = st.slider("Seuil minimum |loading| (PC1)", 0.0, 1.0, 0.3,
                                  step=0.05, key="acp_thresh")
            selected_acp = abs_load[abs_load >= threshold].index.tolist()
            st.info(f"**{len(selected_acp)}** variables avec |loading| ≥ {threshold} sur PC1")

            fig_bar = px.bar(
                x=abs_load.index, y=abs_load.values,
                labels={"x": "Variable", "y": "|Loading PC1|"},
                title="Contributions absolues à PC1",
            )
            fig_bar.add_hline(y=threshold, line_dash="dash", line_color="red")
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.warning("Pas assez de données ou de variables pour l'ACP.")

# Types d'agrégation (smart defaults via get_default_agg_type)
AGG_LABELS = ["Flux (somme)", "Taux (moyenne)", "Stock (dernier)"]
_AGG_LABEL_MAP = {"sum": "Flux (somme)", "mean": "Taux (moyenne)", "last": "Stock (dernier)"}
_AGG_KEY_MAP = {"Flux (somme)": "flow", "Taux (moyenne)": "mean", "Stock (dernier)": "stock"}

with st.expander("⚙️ Types d'agrégation mensuel→trimestriel"):
    agg_types = {}
    cols_per_row = 4
    for i in range(0, len(hf_vars), cols_per_row):
        cols = st.columns(min(cols_per_row, len(hf_vars) - i))
        for j, col in enumerate(cols):
            if i + j < len(hf_vars):
                v = hf_vars[i + j]
                with col:
                    default_label = _AGG_LABEL_MAP[get_default_agg_type(v)]
                    agg_types[v] = st.selectbox(
                        v[:20], AGG_LABELS,
                        index=AGG_LABELS.index(default_label),
                        key=f"agg_{v}",
                    )
    agg_map = {v: _AGG_KEY_MAP[t] for v, t in agg_types.items()}

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
        auto_disagg = st.checkbox("Trimestrialiser automatiquement",
                                  key="now_auto_disagg")
        disagg_method = st.selectbox(
            "Méthode de trimestrialisation",
            ["Chow-Lin (ML)", "Denton-Cholette", "Ecotrim (Fernandez)", "Uniforme"],
            key="now_disagg_method",
        )
        
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
                
                method_key = {
                    "Chow-Lin (ML)": "chow-lin",
                    "Denton-Cholette": "denton",
                    "Ecotrim (Fernandez)": "ecotrim",
                    "Uniforme": None,
                }[disagg_method]
                
                if method_key:
                    pib_q = disaggregate_annual_to_quarterly(
                        pib_annual, hf_q_for_td, method=method_key
                    )
                else:
                    pib_q = disaggregate_annual_to_quarterly(pib_annual)
                st.success(f"✅ PIB trimestrialisé ({disagg_method})")
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
        
        pib_aligned = pib_q.copy()
        
        results = run_nowcast(
            pib_aligned, hf_q,
            models=models,
            h_ahead=h_ahead,
            n_components=n_components,
        )
        
        # Afficher les erreurs d'alignement éventuelles
        _nowcast_errors = results.pop("_errors", [])
        if _nowcast_errors:
            for _e in _nowcast_errors:
                st.warning(f"⚠️ {_e}")
        
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
        "Séries prolongées": "Oui" if st.session_state.get("forecasts", {}).get(country) is not None else "Non",
    }
    # Construire le DataFrame HF mensuel pour l'audit
    _hf_export_df = None
    if hf_df is not None and hf_vars:
        _hf_export_cols = [date_col] + [v for v in hf_vars if v in hf_df.columns]
        _hf_export_df = hf_df[_hf_export_cols].copy().rename(columns={date_col: "Date"})
    data = write_nowcast_excel(
        pib_aligned, results, params,
        hf_vars=hf_vars,
        agg_map=agg_map if "agg_map" in dir() else None,
        hf_df=_hf_export_df,
    )
    download_button(
        data,
        f"RESULT_NOWCAST_{country}_{pd.Timestamp.now().strftime('%Y-%m-%d')}.xlsx",
        "📥 Télécharger le Nowcast",
    )
