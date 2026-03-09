"""Configuration globale de l'application ICAE Streamlit."""
import os
from pathlib import Path

# ── Chemins ────────────────────────────────────────────────────────────────
APP_DIR = Path(__file__).resolve().parent
WORKSPACE = APP_DIR.parent.parent          # .../Mars 2026/
LIVRABLES = WORKSPACE / "Livrables"
CONSOLIDES = LIVRABLES / "01_ICAE_Consolides_Corriges"
CONSOLIDES_FC = LIVRABLES / "01_ICAE_Consolides_Corriges_with_forecast_by_series"
PREVISIONS = LIVRABLES / "02_Previsions_Q1_2026"
ASSETS = APP_DIR / "assets"
LOGO_PATH = ASSETS / "logo_beac.jpg"

# Mode cloud : les dossiers locaux n'existent pas
CLOUD_MODE = not CONSOLIDES.exists()

# ── Pays ───────────────────────────────────────────────────────────────────
COUNTRY_CODES = ["CMR", "CNG", "GAB", "GNQ", "RCA", "TCD"]
COUNTRY_NAMES = {
    "CMR": "Cameroun",
    "CNG": "Congo",
    "GAB": "Gabon",
    "GNQ": "Guinée Équatoriale",
    "RCA": "RCA",
    "TCD": "Tchad",
}

# Poids PIB 2014 (milliards FCFA)
PIB_2014 = {
    "CMR": 13651.4,
    "CNG": 5729.7,
    "GAB": 6861.0,
    "GNQ": 6413.3,
    "RCA": 624.7,
    "TCD": 4989.2,
}
PIB_TOTAL = sum(PIB_2014.values())
POIDS_PIB = {k: v / PIB_TOTAL for k, v in PIB_2014.items()}

# ── Feuilles attendues ────────────────────────────────────────────────────
SHEETS_PAYS = [
    "Consignes", "Codification", "Donnees_calcul",
    "CALCUL_ICAE", "Contrib", "Resultats_Trim",
]
SHEETS_CEMAC = ["Poids_PIB", "ICAE_Pays", "ICAE_Trimestriel", "Consignes"]

# ── Constantes de calcul ──────────────────────────────────────────────────
SYM_FACTOR = 200          # C = 200*(X-X_1)/(X+X_1)
I_INIT = 100              # valeur initiale de l'indice récursif
DENOM_FLOOR = 1e-9        # plancher pour division par zéro dans TCS
SIGMA_FLOOR = 1e-6        # plancher écart-type
SUM_M_CAP = 199.999       # cap sur sum_m pour éviter explosions

# ── Couleurs ──────────────────────────────────────────────────────────────
COLOR_HIST = "#1F4E79"     # bleu foncé – données historiques
COLOR_HIST_BG = "#DCE6F1"  # bleu clair – fond historique
COLOR_FCST = "#E26B0A"     # orange foncé – prévisions
COLOR_FCST_BG = "#FDE9D9"  # orange clair – fond prévisions

# ── Méthodes de prévision ─────────────────────────────────────────────────
FORECAST_METHODS = ["MM3", "MM6", "MM12", "NS", "CS", "TL", "ARIMA", "ETS"]
NAIVE_METHODS = ["MM3", "MM6", "MM12", "NS", "CS", "TL"]

# ── Modèles Nowcast ───────────────────────────────────────────────────────
NOWCAST_MODELS = ["Bridge", "U-MIDAS", "PC", "DFM"]
