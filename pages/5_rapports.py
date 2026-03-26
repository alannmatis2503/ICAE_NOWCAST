"""Module 5 — Rapports Word (Note ICAE + Note Nowcast)."""
import random
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import COUNTRY_CODES, COUNTRY_NAMES, LOGO_PATH
from io_utils.word_report import generate_note_icae, generate_note_nowcast
from ui.charts import chart_quarterly_contrib_ga, fig_to_png_bytes
from ui.components import download_button


# ─────────────────────────────────────────────────────────────────────────
# Synonymes des 4 secteurs (variété dans les rapports)
# ─────────────────────────────────────────────────────────────────────────
SECTOR_SYNONYMS = {
    "Produits de base": [
        "le secteur des produits de base",
        "les activités extractives et primaires",
        "le secteur primaire (pétrole, mines, agriculture)",
        "les industries extractives et produits de base",
        "le segment des matières premières",
    ],
    "Demande privée": [
        "la demande privée",
        "le secteur de la consommation et de l'investissement privés",
        "les activités liées à la demande intérieure privée",
        "le commerce et la demande des ménages",
        "la consommation privée et le BTP",
    ],
    "Secteur public": [
        "le secteur public",
        "les dépenses et activités publiques",
        "la demande publique",
        "les administrations publiques",
        "le secteur gouvernemental",
    ],
    "Financement du secteur privé": [
        "le financement du secteur privé",
        "le crédit à l'économie",
        "les conditions de financement de l'économie",
        "le secteur monétaire et financier",
        "le crédit bancaire au secteur privé",
    ],
}


def _sector_synonym(sector_name: str) -> str:
    """Retourne un synonyme aléatoire du nom de secteur."""
    synonyms = SECTOR_SYNONYMS.get(sector_name)
    if synonyms:
        return random.choice(synonyms)
    return sector_name


# ── Détection de scénario ─────────────────────────────────────────────────
def detect_scenario(ga_trim_series: pd.Series, idx_recent: int,
                    idx_persp: int | None = None) -> dict:
    """
    Détecte le scénario conjoncturel à partir des GA trimestriels.
    idx_recent : position dans la série COMPLÈTE (y compris NaN).
    On utilise les positions originales pour accéder aux valeurs.
    """
    full = ga_trim_series.reset_index(drop=True)
    n = len(full)

    def _val(pos):
        if 0 <= pos < n:
            v = full.iloc[pos]
            return round(float(v), 1) if pd.notna(v) else None
        return None

    current = _val(idx_recent)
    if current is None:
        return {"scenario": "inconnu", "label": "évolution indéterminée",
                "verbe_evo": "évolué", "verbe_persp": "évoluer",
                "ga_current": None, "ga_previous_q": None,
                "ga_year_ago": None, "ga_persp": None}

    prev_q = _val(idx_recent - 1)
    year_ago = _val(idx_recent - 4)

    info = {
        "ga_current": current,
        "ga_previous_q": prev_q,
        "ga_year_ago": year_ago,
    }

    info["ga_persp"] = _val(idx_persp) if idx_persp is not None else None

    if current > 0 and prev_q is not None and prev_q < 0:
        info.update(scenario="reprise", label="une reprise de la croissance",
                    verbe_evo="repris", verbe_persp="se consolider",
                    accroche_adj="rebondi")
    elif current < 0 and prev_q is not None and prev_q > 0:
        info.update(scenario="retournement", label="un retournement à la baisse",
                    verbe_evo="basculé en zone négative",
                    verbe_persp="se redresser", accroche_adj="reculé")
    elif current < 0:
        if prev_q is not None and current < prev_q:
            info.update(scenario="contraction_aggravee",
                        label="un approfondissement de la contraction",
                        verbe_evo="poursuivi son repli",
                        verbe_persp="se stabiliser",
                        accroche_adj="poursuivi son repli")
        else:
            info.update(scenario="contraction",
                        label="une contraction de l'activité",
                        verbe_evo="reculé", verbe_persp="se redresser",
                        accroche_adj="reculé")
    elif prev_q is not None and current > prev_q and (year_ago is None or current > year_ago):
        info.update(scenario="acceleration",
                    label="une accélération de la croissance",
                    verbe_evo="accéléré", verbe_persp="s'accélérer",
                    accroche_adj="accéléré")
    elif prev_q is not None and current < prev_q:
        info.update(scenario="ralentissement",
                    label="un ralentissement de la croissance",
                    verbe_evo="ralenti", verbe_persp="se modérer",
                    accroche_adj="ralenti")
    else:
        info.update(scenario="croissance_stable",
                    label="une croissance stable",
                    verbe_evo="crû", verbe_persp="se maintenir",
                    accroche_adj="progressé")

    return info


SCENARIO_TEMPLATES = {
    "acceleration": {
        "accroche": "Le rythme de progression des activités économiques a accéléré",
        "p1": (
            "a renforcé sa dynamique de croissance. L'ICAE s'est accru "
            "de {ga_current} % en glissement annuel, après {ga_previous_q} % "
            "au trimestre précédent, confirmant une tendance haussière."
        ),
    },
    "ralentissement": {
        "accroche": "Le rythme de progression des activités économiques a ralenti",
        "p1": (
            "a vu son rythme de croissance se modérer. L'ICAE a progressé "
            "de {ga_current} % en glissement annuel, après {ga_previous_q} % "
            "au trimestre précédent, reflétant un essoufflement de la dynamique."
        ),
    },
    "reprise": {
        "accroche": "Les activités économiques ont rebondi",
        "p1": (
            "a renoué avec la croissance. L'ICAE s'est accru "
            "de {ga_current} % en glissement annuel, après une contraction "
            "de {ga_previous_q} % au trimestre précédent, signalant une reprise."
        ),
    },
    "retournement": {
        "accroche": "Les activités économiques ont basculé en zone de contraction",
        "p1": (
            "a vu son activité se retourner. L'ICAE s'est replié "
            "de {ga_current} % en glissement annuel, après une croissance "
            "de {ga_previous_q} % au trimestre précédent."
        ),
    },
    "contraction": {
        "accroche": "Les activités économiques poursuivent leur contraction",
        "p1": (
            "reste en zone de contraction. L'ICAE a reculé "
            "de {ga_current} % en glissement annuel, après {ga_previous_q} % "
            "au trimestre précédent."
        ),
    },
    "contraction_aggravee": {
        "accroche": "La contraction des activités économiques s'est approfondie",
        "p1": (
            "voit sa contraction s'aggraver. L'ICAE a chuté "
            "de {ga_current} % en glissement annuel, après {ga_previous_q} % "
            "au trimestre précédent."
        ),
    },
    "croissance_stable": {
        "accroche": "La croissance des activités économiques s'est maintenue",
        "p1": (
            "a maintenu un rythme de croissance stable. L'ICAE a progressé "
            "de {ga_current} % en glissement annuel, après {ga_previous_q} % "
            "au trimestre précédent."
        ),
    },
}


def _identify_drivers(contrib_row: pd.Series) -> tuple[list, list]:
    """Identifie les secteurs moteurs (+) et freineurs (-)."""
    drivers, drags = [], []
    for sect, val in contrib_row.items():
        if sect in ("trimestre", "Date", "_quarter"):
            continue
        if pd.isna(val):
            continue
        if val > 0.01:
            drivers.append((sect, round(float(val), 2)))
        elif val < -0.01:
            drags.append((sect, round(float(val), 2)))
    drivers.sort(key=lambda x: x[1], reverse=True)
    drags.sort(key=lambda x: x[1])
    return drivers, drags


def _build_sector_commentary(drivers, drags) -> str:
    """Construit un paragraphe décrivant les contributions sectorielles."""
    parts = []
    if drivers:
        names = [_sector_synonym(d[0]) for d in drivers]
        contribs = [f"+{d[1]} %" for d in drivers]
        if len(names) == 1:
            parts.append(f"Cette évolution a été principalement tirée par "
                         f"{names[0]} ({contribs[0]})")
        else:
            parts.append(f"Cette évolution a été principalement tirée par "
                         f"{', '.join(names[:-1])} et {names[-1]} "
                         f"({', '.join(contribs)})")
    if drags:
        names = [_sector_synonym(d[0]) for d in drags]
        contribs = [f"{d[1]} %" for d in drags]
        if parts:
            if len(names) == 1:
                parts.append(f", tandis que {names[0]} ({contribs[0]}) "
                             f"a pesé sur la dynamique.")
            else:
                parts.append(f", tandis que {', '.join(names[:-1])} et "
                             f"{names[-1]} ({', '.join(contribs)}) "
                             f"ont freiné l'activité.")
        else:
            parts.append(f"L'activité a été pénalisée par "
                         f"{', '.join(names)} ({', '.join(contribs)}).")
    if not parts:
        return "Les contributions sectorielles sont restées proches de zéro."
    return "".join(parts) + ("." if not parts[-1].endswith(".") else "")


# ════════════════════════════════════════════════════════════════════════════
# Interface principale
# ════════════════════════════════════════════════════════════════════════════
st.title("📋 Module 5 — Rapports")

report_type = st.radio(
    "Type de rapport",
    ["Note ICAE (CEMAC ou Pays)", "Note Nowcast"],
    horizontal=True, key="report_type",
)

# ════════════════════════════════════════════════════════════════════════════
# NOTE ICAE
# ════════════════════════════════════════════════════════════════════════════
if report_type == "Note ICAE (CEMAC ou Pays)":
    st.header("Note ICAE — Rédaction assistée")

    # ── Source des données ────────────────────────────────────────────────
    data_source = st.radio(
        "Source des données trimestrielles",
        ["Données calculées dans l'application", "Importer un fichier externe"],
        horizontal=True, key="report_data_source",
    )

    q_data = None
    ct_data = None
    ga_trim = None
    entity_name = ""
    entity = "XXX"
    forecast_boundary = None

    if data_source == "Importer un fichier externe":
        st.warning(
            "⚠️ **Format attendu du fichier Excel**\n\n"
            "Le fichier doit contenir une feuille **`Resultats_Trim`** "
            "(ou la première feuille sera utilisée) avec les colonnes :\n\n"
            "| Colonne | Description | Exemple |\n"
            "|---------|-------------|--------|\n"
            "| `Trimestre` | Label du trimestre | 2024T1 |\n"
            "| `ICAE_Trim` | ICAE trimestriel moyen | 102.5 |\n"
            "| `GA_Trim` | Glissement annuel (%) | 3.2 |\n"
            "| `GT_Trim` | Glissement trimestriel (%) | 0.8 |\n\n"
            "**Optionnel** — une feuille **`Contributions`** avec :\n"
            "- `Trimestre` + une colonne par secteur "
            "(Produits de base, Demande privée, Secteur public, "
            "Financement du secteur privé)\n"
            "- Les valeurs doivent être en % (contributions décomposant le GA)."
        )
        uploaded_report = st.file_uploader(
            "Fichier de données trimestrielles (.xlsx)", type=["xlsx"],
            key="report_upload",
        )
        if uploaded_report is None:
            st.info("Veuillez charger un fichier pour continuer.")
            st.stop()

        entity_name = st.text_input(
            "Nom de l'entité (pays ou zone)", value="Pays",
            key="report_entity_name_input",
        )
        entity = entity_name.replace(" ", "_")[:10]

        try:
            xls = pd.ExcelFile(uploaded_report)
            sheets = xls.sheet_names
            target_sheet = sheets[0]
            for s in sheets:
                if "result" in s.lower() or "trim" in s.lower():
                    target_sheet = s
                    break

            q_data = pd.read_excel(uploaded_report, sheet_name=target_sheet)
            uploaded_report.seek(0)

            rename_map = {}
            for c in q_data.columns:
                cl = str(c).lower().strip().replace(" ", "_")
                if "trimestre" in cl or cl == "trim":
                    rename_map[c] = "trimestre"
                elif "icae" in cl and "trim" in cl:
                    rename_map[c] = "icae_trim"
                elif cl.startswith("ga"):
                    rename_map[c] = "GA_Trim"
                elif cl.startswith("gt"):
                    rename_map[c] = "GT_Trim"
            q_data = q_data.rename(columns=rename_map)

            required_cols = ["trimestre", "GA_Trim"]
            missing_cols = [c for c in required_cols if c not in q_data.columns]
            if missing_cols:
                st.error(
                    f"Colonnes manquantes : {', '.join(missing_cols)}. "
                    "Vérifiez la structure du fichier."
                )
                st.stop()

            ct_data = None
            for s in sheets:
                if "contrib" in s.lower():
                    ct_data = pd.read_excel(uploaded_report, sheet_name=s)
                    uploaded_report.seek(0)
                    for c in ct_data.columns:
                        if "trim" in str(c).lower():
                            ct_data = ct_data.rename(columns={c: "trimestre"})
                            break
                    break

            ga_trim = q_data["GA_Trim"]
            st.success(f"✅ {len(q_data)} trimestres chargés.")
        except Exception as e:
            st.error(f"Erreur de lecture du fichier : {e}")
            st.stop()

    else:
        # ── Sélection entité ──────────────────────────────────────────────
        entity_options = ["CEMAC"] + COUNTRY_CODES
        entity = st.selectbox("Entité", entity_options, key="report_entity")
        entity_name = ("LA CEMAC" if entity == "CEMAC"
                       else COUNTRY_NAMES.get(entity, entity))

        if entity == "CEMAC":
            if "cemac_quarterly" in st.session_state:
                q_data = st.session_state["cemac_quarterly"]
                # Harmoniser le nom de la colonne trimestre (CEMAC utilise "Trimestre")
                if "Trimestre" in q_data.columns and "trimestre" not in q_data.columns:
                    q_data = q_data.rename(columns={"Trimestre": "trimestre"})
                # S'assurer que GA_Trim existe
                if "GA_Trim" not in q_data.columns and "ICAE_CEMAC" in q_data.columns:
                    q_data["GA_Trim"] = (
                        q_data["ICAE_CEMAC"] / q_data["ICAE_CEMAC"].shift(4) - 1
                    ) * 100
                if "GT_Trim" not in q_data.columns and "ICAE_CEMAC" in q_data.columns:
                    q_data["GT_Trim"] = (
                        q_data["ICAE_CEMAC"] / q_data["ICAE_CEMAC"].shift(1) - 1
                    ) * 100
                if "icae_trim" not in q_data.columns and "ICAE_CEMAC" in q_data.columns:
                    q_data["icae_trim"] = q_data["ICAE_CEMAC"]
                # Construire les contributions par pays comme proxy de contributions sectorielles
                _poids_used = st.session_state.get("cemac_computed_poids", {})
                if _poids_used:
                    _cemac_ct_cols = []
                    for code in COUNTRY_CODES:
                        if code in q_data.columns:
                            ga_pays = (q_data[code] / q_data[code].shift(4) - 1) * 100
                            col_name = COUNTRY_NAMES.get(code, code)
                            q_data[col_name] = ga_pays * _poids_used.get(code, 0)
                            _cemac_ct_cols.append(col_name)
                    if _cemac_ct_cols and ct_data is None:
                        ct_data = q_data[["trimestre"] + _cemac_ct_cols].copy()
            elif "cemac_result" in st.session_state:
                st.info(
                    "Les résultats CEMAC sont disponibles mais sans détail "
                    "trimestriel complet. Relancez le Module 4 si nécessaire."
                )
        else:
            q_data = (st.session_state.get("icae_quarterly") or {}).get(entity)
            ct_data = (st.session_state.get("icae_contrib_trim") or {}).get(entity)

        if q_data is None:
            st.warning(
                f"⚠️ Pas de données trimestrielles pour {entity_name}. "
                "Calculez l'ICAE dans le Module 1 (ou 4 pour CEMAC) d'abord."
            )
            st.stop()

        ga_trim = q_data["GA_Trim"]
        forecast_boundary = (
            st.session_state.get("icae_forecast_boundary") or {}
        ).get(entity)

    # Harmoniser le nom de la colonne trimestre pour toutes les sources
    if "trimestre" not in q_data.columns:
        for candidate in ["Trimestre", "quarter", "Quarter", "TRIMESTRE"]:
            if candidate in q_data.columns:
                q_data = q_data.rename(columns={candidate: "trimestre"})
                break
        else:
            # Dernière tentative : construire la colonne trimestre
            if q_data.index.dtype == "period[Q-DEC]":
                q_data["trimestre"] = q_data.index.astype(str)
            else:
                st.error("Colonne 'trimestre' introuvable dans les données trimestrielles.")
                st.stop()

    # Construire les listes de trimestres disponibles
    trimestres_list = q_data["trimestre"].astype(str).tolist()

    # ── Choix des trimestres ──────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🎯 Sélection des trimestres d'analyse")
    st.caption("Choisissez le trimestre « évolution récente » (sera comparé au trimestre "
               "précédent et au même trimestre de l'année précédente) et le trimestre "
               "« perspectives à court terme ».")

    # Filtrer les trimestres avec GA non-NaN pour les choix
    valid_indices = [i for i, v in enumerate(ga_trim) if pd.notna(v)]
    valid_trims = [trimestres_list[i] for i in valid_indices]

    if len(valid_trims) < 2:
        st.error("Pas assez de trimestres avec GA calculé pour l'analyse.")
        st.stop()

    # Defaults intelligents basés sur la frontière historique/prévision
    default_recent_idx = max(0, len(valid_trims) - 2)
    default_persp_idx = len(valid_trims) - 1
    if forecast_boundary and forecast_boundary in valid_trims:
        fb_pos = valid_trims.index(forecast_boundary)
        default_recent_idx = max(0, fb_pos - 1)
        default_persp_idx = fb_pos

    col1, col2 = st.columns(2)
    with col1:
        trim_recent_label = st.selectbox(
            "Trimestre — Évolution récente",
            valid_trims,
            index=default_recent_idx,
            key="report_trim_recent",
        )
    with col2:
        trim_persp_label = st.selectbox(
            "Trimestre — Perspectives",
            valid_trims,
            index=default_persp_idx,
            key="report_trim_persp",
        )

    idx_recent = trimestres_list.index(trim_recent_label)
    idx_persp = trimestres_list.index(trim_persp_label)

    # Extraire année et trimestre du label (ex: "2025Q4" ou "2025T4")
    def _parse_trim_label(label):
        s = label.upper().replace("Q", "T")
        for sep in ["T"]:
            if sep in s:
                parts = s.split(sep)
                return int(parts[0]), f"T{parts[1]}"
        return 2025, "T4"

    annee_recent, trim_recent = _parse_trim_label(trim_recent_label)
    annee_persp, trim_persp = _parse_trim_label(trim_persp_label)

    # ── Détection du scénario ─────────────────────────────────────────────
    scenario_info = detect_scenario(ga_trim, idx_recent, idx_persp)
    sc_key = scenario_info.get("scenario", "inconnu")
    ga_c = scenario_info.get("ga_current", "X,X")
    ga_pq = scenario_info.get("ga_previous_q", "Y,Y")
    ga_ya = scenario_info.get("ga_year_ago", "Z,Z")
    ga_persp = scenario_info.get("ga_persp", "W,W")

    st.info(f"🎯 **Scénario détecté** : {scenario_info.get('label', '?')} "
            f"| GA {trim_recent_label} = **{ga_c} %** "
            f"| Trim. précédent = {ga_pq} %"
            f"| Même trim. an-1 = {ga_ya} %")

    # ── Contributions sectorielles pour les trimestres choisis ─────────────
    drivers_recent, drags_recent = [], []
    drivers_persp, drags_persp = [], []
    if ct_data is not None and len(ct_data) > idx_recent:
        row_recent = ct_data.iloc[idx_recent]
        drivers_recent, drags_recent = _identify_drivers(row_recent)
    if ct_data is not None and len(ct_data) > idx_persp:
        row_persp = ct_data.iloc[idx_persp]
        drivers_persp, drags_persp = _identify_drivers(row_persp)

    # Noms des secteurs moteurs pour les templates
    driver_names_str = ", ".join([_sector_synonym(d[0]) for d in drivers_recent]) if drivers_recent else "[les secteurs moteurs]"

    # ══════════════════════════════════════════════════════════════════════
    # Section 1 : Evolution récente
    # ══════════════════════════════════════════════════════════════════════
    st.header(f"📊 Section 1 — Evolution récente ({trim_recent_label})")

    col_chart, col_text = st.columns([1, 1])

    with col_chart:
        st.subheader("Graphique — Contributions et GA")
        st.caption("Définissez la période du graphique ci-dessous.")
        # Sélection de la période du graphique
        c1, c2 = st.columns(2)
        with c1:
            chart_start_evo = st.selectbox(
                "Début", trimestres_list,
                index=max(0, idx_recent - 7),
                key="chart_evo_start",
            )
        with c2:
            chart_end_evo = st.selectbox(
                "Fin", trimestres_list,
                index=idx_recent,
                key="chart_evo_end",
            )
        i_start = trimestres_list.index(chart_start_evo)
        i_end = trimestres_list.index(chart_end_evo) + 1

        if ct_data is not None and len(ct_data) >= i_end:
            sub_ct = ct_data.iloc[i_start:i_end].reset_index(drop=True)
            sub_ga = ga_trim.iloc[i_start:i_end].reset_index(drop=True)
            sub_trims = trimestres_list[i_start:i_end]

            fig_evo = chart_quarterly_contrib_ga(
                sub_ct, sub_ga, sub_trims,
                title=f"ICAE {entity_name} — Contributions sectorielles et GA trimestriel (%)",
            )
            st.plotly_chart(fig_evo, use_container_width=True, key="fig_evo_chart")
            st.session_state["_fig_evo"] = fig_evo
        else:
            st.warning("Contributions sectorielles trimestrielles non disponibles.")

        # Tableau des valeurs
        with st.expander("📋 Valeurs trimestrielles"):
            display_df = q_data.iloc[i_start:i_end][
                ["trimestre", "icae_trim", "GA_Trim", "GT_Trim"]
            ].copy()
            for c in ("icae_trim", "GA_Trim", "GT_Trim"):
                display_df[c] = display_df[c].round(2)
            st.dataframe(display_df, use_container_width=True, hide_index=True)

    with col_text:
        st.subheader("Texte — Évolution récente")

        tpl = SCENARIO_TEMPLATES.get(sc_key, SCENARIO_TEMPLATES["croissance_stable"])

        default_accroche = f"{tpl['accroche']} au {trim_recent} {annee_recent}."
        accroche_evo = st.text_area(
            "Accroche (sous-titre en gras)", value=default_accroche,
            height=80, key="accroche_evo",
        )

        p1_scenario = tpl["p1"].format(ga_current=ga_c, ga_previous_q=ga_pq)
        sector_commentary = _build_sector_commentary(drivers_recent, drags_recent)

        ya_comparison = ""
        if ga_ya is not None:
            ya_comparison = (f" En comparaison avec le {trim_recent} {annee_recent - 1}, "
                             f"où le GA était de {ga_ya} %, l'évolution témoigne de "
                             f"{scenario_info.get('label', 'un changement de dynamique')}.")

        default_paras = [
            f"Au cours du {trim_recent} {annee_recent}, le secteur productif de "
            f"{entity_name} {p1_scenario}{ya_comparison}",
            sector_commentary,
            f"La dynamique de croissance a ainsi été soutenue par {driver_names_str}.",
        ]

        paras_evo = []
        for i, dp in enumerate(default_paras):
            p = st.text_area(f"Paragraphe {i + 1}", value=dp, height=100,
                             key=f"para_evo_{i}")
            paras_evo.append(p)

    st.markdown("---")

    # ══════════════════════════════════════════════════════════════════════
    # Section 2 : Perspectives à court terme
    # ══════════════════════════════════════════════════════════════════════
    st.header(f"🔮 Section 2 — Perspectives ({trim_persp_label})")

    col_chart2, col_text2 = st.columns([1, 1])

    with col_chart2:
        st.subheader("Graphique — Avec projection")
        st.caption("Même graphique que l'évolution récente, mais étendu "
                   "au trimestre de perspective (pointillés).")
        c3, c4 = st.columns(2)
        with c3:
            chart_start_persp = st.selectbox(
                "Début", trimestres_list,
                index=max(0, idx_persp - 7),
                key="chart_persp_start",
            )
        with c4:
            chart_end_persp = st.selectbox(
                "Fin", trimestres_list,
                index=min(len(trimestres_list) - 1, idx_persp),
                key="chart_persp_end",
            )
        j_start = trimestres_list.index(chart_start_persp)
        j_end = trimestres_list.index(chart_end_persp) + 1

        if ct_data is not None and len(ct_data) >= j_end:
            sub_ct2 = ct_data.iloc[j_start:j_end].reset_index(drop=True)
            sub_ga2 = ga_trim.iloc[j_start:j_end].reset_index(drop=True)
            sub_trims2 = trimestres_list[j_start:j_end]

            # Déterminer l'index de début des pointillés (forecast)
            # = position du trimestre perspective dans la sous-série
            fcst_idx = idx_persp - j_start if idx_persp >= j_start else None

            fig_persp = chart_quarterly_contrib_ga(
                sub_ct2, sub_ga2, sub_trims2,
                title=f"ICAE {entity_name} — Perspectives {trim_persp_label}",
                forecast_start_idx=fcst_idx,
            )
            st.plotly_chart(fig_persp, use_container_width=True, key="fig_persp_chart")
            st.session_state["_fig_persp"] = fig_persp
        else:
            st.warning("Pas de contributions pour la période de perspectives.")

    with col_text2:
        st.subheader("Texte — Perspectives")

        driver_names_persp = ", ".join([_sector_synonym(d[0]) for d in drivers_persp]) if drivers_persp else "[les secteurs moteurs]"

        if sc_key in ("ralentissement", "contraction", "contraction_aggravee",
                      "retournement"):
            persp_intro = (
                f"Les indicateurs avancés suggèrent un redressement progressif "
                f"de l'activité au {trim_persp} {annee_persp}, sous réserve d'une amélioration du contexte "
                f"international et d'un maintien des prix des matières premières."
            )
        elif sc_key in ("acceleration", "reprise"):
            persp_intro = (
                f"La dynamique de croissance devrait se poursuivre au {trim_persp} {annee_persp}, "
                f"soutenue par {driver_names_persp}. Toutefois, un ralentissement "
                f"n'est pas exclu si les conditions extérieures se détériorent."
            )
        else:
            persp_intro = (
                f"Les projections indiquent un maintien du rythme de croissance "
                f"actuel au {trim_persp} {annee_persp}, soutenu par {driver_names_persp}."
            )

        sector_commentary_persp = _build_sector_commentary(drivers_persp, drags_persp)

        ga_persp_str = ga_persp if ga_persp is not None else "[X,X]"
        default_persp = [
            persp_intro,
            sector_commentary_persp,
            f"En synthèse, l'ICAE de {entity_name} devrait progresser de "
            f"{ga_persp_str} % en GA au {trim_persp} {annee_persp} "
            f"(après {ga_c} % au {trim_recent} {annee_recent})"
            f"{f', contre {ga_ya} % au {trim_recent} {annee_recent - 1}' if ga_ya is not None else ''}.",
        ]

        accroche_persp = st.text_area(
            "Accroche perspectives",
            value=f"Le rythme de progression des activités économiques "
                  f"devrait {scenario_info.get('verbe_persp', 'évoluer')} au "
                  f"{trim_persp} {annee_persp}.",
            height=80, key="accroche_persp",
        )

        paras_persp = []
        for i, dp in enumerate(default_persp):
            p = st.text_area(f"Paragraphe perspectives {i + 1}", value=dp,
                             height=100, key=f"para_persp_{i}")
            paras_persp.append(p)

    st.markdown("---")

    # ── Génération du document ────────────────────────────────────────────
    st.header("📄 Génération du document Word")

    if st.button("📝 Générer la Note ICAE", type="primary", key="gen_note_icae"):
        # Convertir les graphiques en PNG
        chart_evo_bytes = fig_to_png_bytes(st.session_state.get("_fig_evo"))
        chart_persp_bytes = fig_to_png_bytes(st.session_state.get("_fig_persp"))

        user_texts = {
            "accroche_evolution": accroche_evo,
            "paragraphs_evolution": paras_evo,
            "accroche_perspectives": accroche_persp,
            "paragraphs_perspectives": paras_persp,
        }

        logo = str(LOGO_PATH) if LOGO_PATH.exists() else None

        doc_bytes = generate_note_icae(
            country_name=entity_name,
            trimestre=trim_recent,
            annee=annee_recent,
            icae_data={},
            chart_evolution_bytes=chart_evo_bytes,
            chart_perspectives_bytes=chart_persp_bytes,
            logo_path=logo,
            user_texts=user_texts,
        )

        download_button(
            doc_bytes,
            f"Note_ICAE_{entity}_{trim_recent}_{annee_recent}.docx",
            "📥 Télécharger la Note ICAE",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.success("✅ Document généré !")


# ════════════════════════════════════════════════════════════════════════════
# NOTE NOWCAST
# ════════════════════════════════════════════════════════════════════════════
elif report_type == "Note Nowcast":
    st.header("Note Nowcast — Rapport technique")

    has_nowcast = "nowcast_results" in st.session_state

    if not has_nowcast:
        st.warning("⚠️ Aucun résultat Nowcast disponible. Lancez le Module 3 d'abord.")
        st.stop()

    nowcast_data = st.session_state["nowcast_results"]

    results_for_report = {}
    chart_bytes = {}

    for country_code, results in nowcast_data.items():
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
            })
        results_for_report[country_code] = {
            "metrics_df": pd.DataFrame(perf_rows),
        }

        st.subheader(f"Résultats — {COUNTRY_NAMES.get(country_code, country_code)}")
        st.dataframe(pd.DataFrame(perf_rows), use_container_width=True,
                     hide_index=True)

    if st.button("📝 Générer la Note Nowcast", type="primary",
                 key="gen_note_nowcast"):
        logo = str(LOGO_PATH) if LOGO_PATH.exists() else None

        doc_bytes = generate_note_nowcast(
            results_by_country=results_for_report,
            chart_bytes=chart_bytes,
            logo_path=logo,
        )

        download_button(
            doc_bytes,
            f"Note_Nowcast_{pd.Timestamp.now().strftime('%Y-%m-%d')}.docx",
            "📥 Télécharger la Note Nowcast",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.success("✅ Document généré !")
