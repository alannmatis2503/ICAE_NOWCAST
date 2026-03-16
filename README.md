---
title: ICAE NOWCAST — BEAC
emoji: 📊
colorFrom: blue
colorTo: indigo
sdk: streamlit
sdk_version: "1.42.0"
app_file: app.py
pinned: false
license: mit
---

# ICAE NOWCAST — BEAC

Application Streamlit pour le calcul de l'**Indice Composite des Activités Économiques (ICAE)** des pays de la CEMAC, développée pour la Banque des États de l'Afrique Centrale (BEAC).

## Fonctionnalités

| Module | Description |
|--------|-------------|
| **Calcul ICAE** | Calcul de l'ICAE à partir de fichiers consolidés Excel |
| **Prévisions** | Prévision des variables composantes (ARIMA, ETS, moyennes mobiles, etc.) |
| **Nowcast** | Estimation en temps réel du PIB (Bridge, U-MIDAS, DFM, PC) |
| **Agrégation CEMAC** | Agrégation pondérée des ICAE des 6 pays |
| **Rapports** | Génération automatisée de notes Word avec graphiques |
| **Documentation** | Guide d'utilisation intégré |

## Installation locale

```bash
# 1. Cloner le dépôt
git clone https://github.com/alannmatis2503/ICAE_NOWCAST.git
cd ICAE_NOWCAST

# 2. Créer un environnement virtuel
python -m venv .venv

# 3. Activer (Windows PowerShell)
.\.venv\Scripts\Activate.ps1
# Ou sous Linux/Mac :
# source .venv/bin/activate

# 4. Installer les dépendances
pip install -r requirements.txt

# 5. Lancer l'application
streamlit run app.py
```

L'application s'ouvre à l'adresse http://localhost:8501.

## Déploiement en ligne

- **GitHub** : https://github.com/alannmatis2503/ICAE_NOWCAST
- **Hugging Face Spaces** : https://huggingface.co/spaces/Born237/ICAE_NOWCAST

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

Usage interne BEAC — MIT License.
