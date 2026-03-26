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
LIVRABLE_FINAL = WORKSPACE / "Livrable_Final" / "01_Classeurs_ICAE"
CEMAC_TEMPLATE = LIVRABLE_FINAL / "ICAE_CEMAC_Consolide.xlsx"
ASSETS = APP_DIR / "assets"
LOGO_PATH = ASSETS / "logo_beac.jpg"

# ── Pays (ordre BEAC : CMR, RCA, CNG, GAB, GNQ, TCD) ──────────────────────
COUNTRY_CODES = ["CMR", "RCA", "CNG", "GAB", "GNQ", "TCD"]
COUNTRY_NAMES = {
    "CMR": "Cameroun",
    "RCA": "RCA",
    "CNG": "Congo",
    "GAB": "Gabon",
    "GNQ": "Guinée Équatoriale",
    "TCD": "Tchad",
}

# Poids PIB 2014 (milliards FCFA) — source : ICAE_CEMAC_Consolide.xlsx
PIB_2014 = {
    "CMR": 22310.90,
    "RCA": 830.35,
    "CNG": 4346.14,
    "GAB": 3297.98,
    "GNQ": 5046.69,
    "TCD": 8637.90,
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

# ── Mode cloud ────────────────────────────────────────────────────────────
CLOUD_MODE = False

# ── Modèles Nowcast ───────────────────────────────────────────────────────
NOWCAST_MODELS = ["Bridge", "U-MIDAS", "PC", "DFM"]

# ── Types d'agrégation par défaut pour les variables connues ──────────────
# "sum" = Flux (somme), "mean" = Taux/Indice (moyenne), "last" = Stock (fin de période)
# Clefs en minuscules partielles pour le matching par label.
_STOCK_KEYWORDS = [
    "monnaie", "m2", "masse monétaire", "masse mon",
    "circulation fiduciaire", "crédit à l'économie", "crédit a l'economie",
    "créances nettes", "creances_nettes", "crédit",
    "inverse créances", "inverse_creances", "inverse du taux",
]
_RATE_KEYWORDS = [
    "taux", "inverse", "ipc", "variation m2", "ratio",
    "créances douteuses", "créances en souffrance",
    "creances_douteuses", "taux_creances",
]

def get_default_agg_type(var_name: str) -> str:
    """Retourne le type d'agrégation par défaut pour une variable.

    Returns 'sum' (flux), 'mean' (taux/indice) ou 'last' (stock).
    """
    vl = var_name.lower()
    for kw in _STOCK_KEYWORDS:
        if kw in vl:
            return "last"
    for kw in _RATE_KEYWORDS:
        if kw in vl:
            return "mean"
    return "sum"  # défaut = flux
