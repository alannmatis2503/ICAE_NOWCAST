---
title: Suivi du secteur productif infra-annuel ICAE et PIB Nowcast
emoji: 📊
colorFrom: blue
colorTo: indigo
sdk: streamlit
sdk_version: "1.41.0"
app_file: app.py
pinned: false
license: other
---

# Suivi du secteur productif infra-annuel — ICAE et PIB Nowcast

Application Streamlit développée pour la **Banque des États de l'Afrique Centrale (BEAC)** — DERSRI.

Elle calcule l'**Indice Composite d'Activité Économique (ICAE)** des pays de la CEMAC et estime le **PIB en temps réel** (Nowcasting).

## Modules

| Module | Description |
|--------|-------------|
| **1 — Calcul ICAE** | Calcul de l'ICAE à partir de fichiers consolidés Excel |
| **2 — Prévisions** | Prévision des variables composantes (ARIMA, ETS, moyennes mobiles, etc.) |
| **3 — Nowcast** | Estimation en temps réel du PIB (Bridge, U-MIDAS, DFM, PC) |
| **4 — Agrégation CEMAC** | Agrégation pondérée des ICAE et PIB Nowcast des 6 pays |
| **5 — Rapports** | Génération automatisée de notes Word avec graphiques |
| **Documentation** | Guide d'utilisation intégré |

## Installation locale

```bash
# 1. Cloner le dépôt
git clone https://github.com/alannmatis2503/ICAE_NOWCAST.git
cd ICAE_NOWCAST

# 2. Accéder au dossier de l'application
cd Livrable_Final/04_Application/ICAE_Streamlit

# 3. Créer et activer un environnement virtuel (Windows)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 4. Installer les dépendances
pip install -r requirements.txt

# 5. Lancer l'application
streamlit run app.py
```

L'application s'ouvre à l'adresse http://localhost:8501.

## Structure du projet

```
app.py              # Point d'entrée Streamlit
config.py           # Configuration (chemins, constantes, couleurs)
requirements.txt    # Dépendances Python
assets/             # Logo et images
core/               # Moteurs de calcul (ICAE, prévisions, nowcast)
io_utils/           # Lecture/écriture Excel et Word
pages/              # Modules de l'application (0 à 5)
ui/                 # Composants graphiques et styles
tests/              # Tests unitaires
```

## Technologies

- **Python 3.12** / **Streamlit** — Interface web interactive
- **Plotly** — Graphiques interactifs
- **openpyxl** / **xlsxwriter** — Manipulation Excel
- **python-docx** — Génération de rapports Word
- **statsmodels** / **pmdarima** — Modèles de séries temporelles
- **scikit-learn** — Machine learning (nowcast)

## Licence

Usage interne BEAC. Tous droits réservés.
