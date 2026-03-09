"""Module 5 — Rapports Word (Note ICAE + Note Nowcast)."""
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

from config import COUNTRY_CODES, COUNTRY_NAMES, LOGO_PATH
from io_utils.word_report import generate_note_icae, generate_note_nowcast
from ui.charts import chart_ga_bars, fig_to_png_bytes
from ui.components import download_button

st.title("📋 Module 5 — Rapports")

report_type = st.radio(
    "Type de rapport",
    ["Note ICAE (CEMAC ou Pays)", "Note Nowcast"],
    horizontal=True,
    key="report_type",
)

# ════════════════════════════════════════════════════════════════════════════
# NOTE ICAE
# ════════════════════════════════════════════════════════════════════════════
if report_type == "Note ICAE (CEMAC ou Pays)":
    st.header("Note ICAE — Rédaction assistée")
    
    # Sélection entité
    entity_options = ["CEMAC"] + COUNTRY_CODES
    entity = st.selectbox("Entité", entity_options, key="report_entity")
    entity_name = "LA CEMAC" if entity == "CEMAC" else COUNTRY_NAMES.get(entity, entity)
    
    col1, col2 = st.columns(2)
    with col1:
        trimestre = st.selectbox("Trimestre", ["T1", "T2", "T3", "T4"],
                                 key="report_trim")
    with col2:
        annee = st.number_input("Année", 2020, 2030, 2025, key="report_year")
    
    # ── Vérifier les données disponibles ──────────────────────────────────
    has_icae = "icae_monthly" in st.session_state
    has_cemac = "cemac_result" in st.session_state
    has_quarterly = "icae_quarterly" in st.session_state
    
    # Construire le graphique de GA pour aide à la rédaction
    chart_evo_bytes = None
    chart_persp_bytes = None
    ga_data = None
    icae_data_display = None
    
    if entity == "CEMAC" and has_cemac:
        result_df = st.session_state["cemac_result"]
        ga_data = result_df["GA"]
        icae_data_display = result_df["ICAE_CEMAC"]
    elif entity in (st.session_state.get("icae_monthly") or {}):
        icae_series = st.session_state["icae_monthly"][entity]
        ga_data = (icae_series / icae_series.shift(12) - 1) * 100
        icae_data_display = icae_series
    
    st.markdown("---")
    st.markdown("""
    ### Mode de rédaction
    
    Chaque section ci-dessous affiche :
    - **Le graphique / tableau interactif** correspondant pour vérifier les données
    - **Un champ de texte pré-rempli** avec un squelette à compléter
    
    Remplissez les textes en vous appuyant sur les graphiques, puis générez le document Word.
    """)
    
    # ══════════════════════════════════════════════════════════════════════
    # Section 1 : Evolution récente
    # ══════════════════════════════════════════════════════════════════════
    st.header("📊 Section 1 — Evolution récente")
    
    # Graphique GA interactif
    if ga_data is not None:
        col_chart, col_text = st.columns([1, 1])
        
        with col_chart:
            st.subheader("Graphique de référence")
            fig_evo = chart_ga_bars(
                ga_data.index, ga_data.values,
                title=f"ICAE {entity_name} — GA (%)",
            )
            st.plotly_chart(fig_evo, use_container_width=True)
            
            # Tableau des dernières valeurs
            st.subheader("Dernières valeurs")
            if icae_data_display is not None:
                recent = pd.DataFrame({
                    "ICAE": icae_data_display.tail(12).values,
                    "GA (%)": ga_data.tail(12).values,
                }, index=ga_data.tail(12).index)
                st.dataframe(recent, use_container_width=True)
            
            # Convertir en PNG pour le Word
            chart_evo_bytes = fig_to_png_bytes(fig_evo)
        
        with col_text:
            st.subheader("Texte à rédiger")
            
            # Accroche
            default_accroche = (
                f"Le rythme de progression des activités économiques a "
                f"[accéléré/ralenti/stagné] au {trimestre} {annee}."
            )
            accroche_evo = st.text_area(
                "Accroche (sous-titre en gras)",
                value=default_accroche,
                height=80,
                key="accroche_evo",
            )
            
            # Paragraphes
            default_paras = [
                f"Au cours du {trimestre} {annee}, le secteur productif de "
                f"{entity_name} a [montré des signes de...]. L'ICAE de "
                f"{entity_name} s'est [accru/replié] de [X,X] % en "
                f"glissement annuel, après [Y,Y] % au trimestre précédent.",
                
                "Les activités extractives (pétrole, manganèse, or...) ont "
                "[contribué positivement/négativement] à cette évolution, "
                "avec [détails sur les volumes/prix].",
                
                "L'industrie manufacturière a [progressé/reculé], tandis que "
                "le secteur du BTP a [détails]. La sylviculture a "
                "[maintenu/amélioré] sa dynamique.",
                
                "Le secteur des transports et services a [complété/freiné] "
                "la dynamique globale.",
            ]
            
            paras_evo = []
            for i, dp in enumerate(default_paras):
                p = st.text_area(
                    f"Paragraphe {i+1}",
                    value=dp,
                    height=100,
                    key=f"para_evo_{i}",
                )
                paras_evo.append(p)
    else:
        st.warning("⚠️ Pas de données ICAE disponibles. Calculez l'ICAE dans "
                   "le Module 1 ou 4 d'abord.")
        accroche_evo = st.text_area("Accroche", key="accroche_evo_fallback",
                                    height=80)
        paras_evo = [st.text_area(f"Paragraphe {i+1}", key=f"para_evo_fb_{i}",
                                  height=100)
                     for i in range(4)]
    
    st.markdown("---")
    
    # ══════════════════════════════════════════════════════════════════════
    # Section 2 : Perspectives
    # ══════════════════════════════════════════════════════════════════════
    st.header("🔮 Section 2 — Perspectives à court terme")
    
    # Graphique avec prévision
    has_forecasts = "forecasts" in st.session_state or "prev_results" in st.session_state
    
    if ga_data is not None:
        col_chart2, col_text2 = st.columns([1, 1])
        
        with col_chart2:
            st.subheader("Graphique avec projection")
            
            if has_forecasts:
                # Inclure les prévisions dans le graphique
                fig_persp = chart_ga_bars(
                    ga_data.index, ga_data.values,
                    title=f"ICAE {entity_name} — GA (%) avec projection",
                    fcst_start=ga_data.index[-4] if len(ga_data) > 4 else None,
                )
            else:
                fig_persp = chart_ga_bars(
                    ga_data.index, ga_data.values,
                    title=f"ICAE {entity_name} — GA (%) [sans projection]",
                )
            
            st.plotly_chart(fig_persp, use_container_width=True)
            chart_persp_bytes = fig_to_png_bytes(fig_persp)
            
            if not has_forecasts:
                st.info("💡 Pour inclure les projections, lancez le Module 2 "
                        "(Prévisions) d'abord.")
        
        with col_text2:
            st.subheader("Texte à rédiger")
            
            # Trimestre suivant
            trim_map = {"T1": "T2", "T2": "T3", "T3": "T4", "T4": "T1"}
            trim_suiv = trim_map[trimestre]
            annee_suiv = annee + 1 if trimestre == "T4" else annee
            
            accroche_persp = st.text_area(
                "Accroche perspectives",
                value=f"Le rythme de progression des activités économiques "
                      f"devrait [s'accélérer/se modérer] au "
                      f"{trim_suiv} {annee_suiv}.",
                height=80,
                key="accroche_persp",
            )
            
            default_persp = [
                "Les projections sectorielles indiquent que [détails par secteur]. "
                "Le secteur pétrolier devrait [maintenir/voir croître/diminuer] "
                "ses volumes.",
                
                "Les facteurs de risque incluent [tensions géopolitiques, "
                "fluctuations des cours des matières premières, conditions "
                "climatiques, etc.].",
                
                f"En synthèse, l'ICAE de {entity_name} devrait progresser de "
                f"[X,X] % en GA (après [Y,Y] % un trimestre plus tôt), "
                f"contre [Z,Z] % un an auparavant.",
            ]
            
            paras_persp = []
            for i, dp in enumerate(default_persp):
                p = st.text_area(
                    f"Paragraphe perspectives {i+1}",
                    value=dp,
                    height=100,
                    key=f"para_persp_{i}",
                )
                paras_persp.append(p)
    else:
        accroche_persp = st.text_area("Accroche perspectives",
                                      key="accroche_persp_fb", height=80)
        paras_persp = [st.text_area(f"Paragraphe perspectives {i+1}",
                                    key=f"para_persp_fb_{i}", height=100)
                       for i in range(3)]
    
    st.markdown("---")
    
    # ── Génération du document ────────────────────────────────────────────
    st.header("📄 Génération du document Word")
    
    if st.button("📝 Générer la Note ICAE", type="primary", key="gen_note_icae"):
        user_texts = {
            "accroche_evolution": accroche_evo,
            "paragraphs_evolution": paras_evo,
            "accroche_perspectives": accroche_persp,
            "paragraphs_perspectives": paras_persp,
        }
        
        logo = str(LOGO_PATH) if LOGO_PATH.exists() else None
        
        doc_bytes = generate_note_icae(
            country_name=entity_name,
            trimestre=trimestre,
            annee=annee,
            icae_data={},
            chart_evolution_bytes=chart_evo_bytes,
            chart_perspectives_bytes=chart_persp_bytes,
            logo_path=logo,
            user_texts=user_texts,
        )
        
        download_button(
            doc_bytes,
            f"Note_ICAE_{entity}_{trimestre}_{annee}.docx",
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
    
    # Préparer les données pour le rapport
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
        
        # Afficher les résultats
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
