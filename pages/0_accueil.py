"""Page d'accueil — BEAC / ICAE."""
import streamlit as st
from config import LOGO_PATH

# ── En-tête ───────────────────────────────────────────────────────────────────
col_l, col_c, col_r = st.columns([1, 3, 1])
with col_c:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=220)

    st.markdown(
        """
        <div style="text-align:center; padding: 0.5rem 0 1.2rem 0;">
            <h1 style="color:#1F4E79; font-size:1.75rem; margin-bottom:0.2rem; line-height:1.3;">
                Suivi du secteur productif infra-annuel
            </h1>
            <h2 style="color:#2E75B6; font-size:1.25rem; font-weight:500; margin:0 0 0.6rem 0;">
                ICAE &amp; PIB Nowcast
            </h2>
            <p style="color:#555; font-size:0.9rem; margin:0;">
                BEAC &mdash; Direction des &Eacute;tudes, de la Recherche,<br>
                des Statistiques et des Relations Internationales (DERSRI)
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ── Présentation des modules ───────────────────────────────────────────────────
st.markdown(
    "<h3 style='color:#1F4E79; margin-bottom:0.8rem;'>Modules disponibles</h3>",
    unsafe_allow_html=True,
)

MODULES = [
    {
        "icon": "📊",
        "num": "Module 1",
        "title": "Calcul ICAE",
        "desc": (
            "Charge un fichier consolidé pays, calcule l'ICAE mensuel et trimestriel "
            "avec normalisation base 100, et exporte les résultats au format Excel."
        ),
        "color": "#1F4E79",
        "bg": "#EBF2FA",
    },
    {
        "icon": "📈",
        "num": "Module 2",
        "title": "Prévisions",
        "desc": (
            "Prolonge les séries temporelles sur un horizon configurable via 8 méthodes "
            "(MM3/6/12, Naïve saisonnière, Croissance saisonnière, Tendance linéaire, ARIMA, ETS). "
            "Sélection automatique par MAPE."
        ),
        "color": "#375623",
        "bg": "#EBF5E8",
    },
    {
        "icon": "🔮",
        "num": "Module 3",
        "title": "Nowcast PIB",
        "desc": (
            "Estime le PIB trimestriel en temps réel à partir d'indicateurs haute fréquence "
            "via quatre modèles : Bridge, U-MIDAS, Composantes Principales (PC) et DFM."
        ),
        "color": "#4A235A",
        "bg": "#F5EBF9",
    },
    {
        "icon": "🌍",
        "num": "Module 4",
        "title": "Agrégation CEMAC",
        "desc": (
            "Agrège les ICAE et les PIB Nowcast des 6 pays de la CEMAC en un indice régional "
            "pondéré par les PIB. Visualisation mensuelle et trimestrielle par pays."
        ),
        "color": "#7B3800",
        "bg": "#FDF2E9",
    },
    {
        "icon": "📋",
        "num": "Module 5",
        "title": "Rapports",
        "desc": (
            "Génère automatiquement une Note conjoncturelle Word illustrée "
            "(graphiques ICAE, GA, contributions sectorielles, Nowcast) "
            "avec détection du scénario conjoncturel."
        ),
        "color": "#7B0000",
        "bg": "#FDEAEA",
    },
]

cols = st.columns(len(MODULES))
for col, m in zip(cols, MODULES):
    with col:
        st.markdown(
            f"""
            <div style="
                background:{m['bg']};
                border-left: 4px solid {m['color']};
                border-radius: 6px;
                padding: 0.85rem 0.9rem;
                height: 100%;
                min-height: 200px;
            ">
                <div style="font-size:1.6rem; margin-bottom:0.3rem;">{m['icon']}</div>
                <div style="color:{m['color']}; font-size:0.7rem; font-weight:700;
                            text-transform:uppercase; letter-spacing:0.05em;">
                    {m['num']}
                </div>
                <div style="color:{m['color']}; font-size:1rem; font-weight:700;
                            margin-bottom:0.5rem;">
                    {m['title']}
                </div>
                <div style="color:#333; font-size:0.82rem; line-height:1.45;">
                    {m['desc']}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.divider()

# ── Workflow recommandé ────────────────────────────────────────────────────────
st.markdown(
    "<h3 style='color:#1F4E79; margin-bottom:0.5rem;'>Utilisation recommandée</h3>",
    unsafe_allow_html=True,
)

steps = [
    ("1", "#1F4E79", "Charger le fichier consolidé pays dans le **Module 1** pour calculer l'ICAE."),
    ("2", "#375623", "Aller dans le **Module 2** pour prolonger les séries et valider les prévisions."),
    ("3", "#7B3800", "Utiliser le **Module 3** pour estimer le PIB trimestriel en temps réel."),
    ("4", "#4A235A", "Agréger les résultats des 6 pays dans le **Module 4** (ICAE CEMAC)."),
    ("5", "#7B0000", "Générer la Note conjoncturelle Word depuis le **Module 5**."),
]

for num, color, text in steps:
    st.markdown(
        f"""
        <div style="display:flex; align-items:flex-start; gap:0.75rem; margin-bottom:0.55rem;">
            <div style="
                background:{color}; color:white;
                border-radius:50%; width:26px; height:26px;
                display:flex; align-items:center; justify-content:center;
                font-size:0.78rem; font-weight:700; flex-shrink:0; margin-top:1px;
            ">{num}</div>
            <div style="color:#333; font-size:0.9rem; padding-top:3px;">{text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.divider()

st.caption(
    "Pour les formules mathématiques, la structure des fichiers attendus et la description "
    "détaillée des méthodes, consultez la page **Documentation / Aide** dans le menu."
)
