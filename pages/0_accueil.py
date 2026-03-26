"""Page d'accueil — BEAC / ICAE."""
import streamlit as st
from config import LOGO_PATH

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=300)
    st.markdown(
        "<h1 style='text-align:center; color:#1F4E79;'>"
        "Indice Composite d'Activité Économique</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<h3 style='text-align:center; color:#2E75B6;'>"
        "BEAC — Direction des Études, de la Recherche, des Statistiques "
        "et des Relations Internationales (DERSRI)</h3>",
        unsafe_allow_html=True,
    )

st.markdown("---")

st.markdown("""
### Bienvenue sur l'application ICAE

Cette application permet de :

| Module | Description |
|--------|-------------|
| **Module 1 — Calcul ICAE** | Calcul de l'ICAE à partir du fichier consolidé pays |
| **Module 2 — Prévisions** | Prolongation des séries par 8 méthodes (MM3/6/12, NS, CS, TL, ARIMA, ETS) |
| **Module 3 — Nowcast** | Estimation du PIB en temps réel (Bridge, U-MIDAS, PC, DFM) |
| **Module 4 — Agrégation CEMAC** | Agrégation pondérée de l'ICAE et du PIB Nowcast des 6 pays de la CEMAC |
| **Module 5 — Rapports** | Génération automatisée de Notes Word (ICAE + Nowcast) |

---

**Utilisation recommandée :**

1. Commencez par le **Module 1** pour charger un fichier consolidé ICAE et calculer l'indicateur.
2. Passez au **Module 2** pour effectuer des prévisions sur les séries individuelles.
3. Après validation des prévisions, l'ICAE est recalculé automatiquement et vous pouvez accéder au **Module 5** pour la rédaction de la Note.
4. Le **Module 3** (Nowcast) permet d'estimer le PIB trimestriel en temps réel.
5. Le **Module 4** agrège les ICAE des 6 pays de la CEMAC.

Consultez la page **Documentation / Aide** pour plus de détails sur les formules et la structure attendue des fichiers.
""")
