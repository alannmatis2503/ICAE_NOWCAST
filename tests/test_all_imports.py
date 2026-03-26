"""Test d'importation de tous les modules modifiés."""
import sys
import os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

mods = [
    ("config", "from config import COUNTRY_CODES, POIDS_PIB, CONSOLIDES"),
    ("icae_engine", "from core.icae_engine import run_icae_pipeline"),
    ("forecast_engine", "from core.forecast_engine import run_all_forecasts"),
    ("nowcast_engine", "from core.nowcast_engine import run_nowcast"),
    ("quarterly", "from core.quarterly import quarterly_mean, calc_ga_trim"),
    ("tempdisagg", "from core.tempdisagg import chow_lin, denton_cholette, fernandez, disaggregate_annual"),
    ("cemac_engine", "from core.cemac_engine import compute_icae_cemac"),
    ("excel_reader", "from io_utils.excel_reader import load_country_file, read_codification"),
    ("excel_writer", "from io_utils.excel_writer import write_previsions_excel"),
    ("charts", "from ui.charts import chart_forecast_comparison, chart_nowcast"),
    ("components", "from ui.components import download_button"),
]

errors = []
for name, imp in mods:
    try:
        exec(imp)
        print(f"OK {name}")
    except Exception as e:
        errors.append(f"{name}: {e}")
        print(f"ERREUR {name}: {e}")

if errors:
    print(f"\n{len(errors)} erreur(s)")
    sys.exit(1)
else:
    print("\nTous les modules importes sans erreur")
