"""Test d'importation de tous les modules."""
import sys
import os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.path.insert(0, r"c:\Users\HP\Documents\Stage pro BEAC\Work\ICAE\Mars 2026\Apps\ICAE_Streamlit")

errors = []

try:
    from config import COUNTRY_CODES, POIDS_PIB, CONSOLIDES
    print("✅ config OK")
except Exception as e:
    errors.append(f"config: {e}")
    print(f"❌ config: {e}")

try:
    from core.icae_engine import run_icae_pipeline
    print("✅ icae_engine OK")
except Exception as e:
    errors.append(f"icae_engine: {e}")
    print(f"❌ icae_engine: {e}")

try:
    from core.forecast_engine import run_all_forecasts
    print("✅ forecast_engine OK")
except Exception as e:
    errors.append(f"forecast_engine: {e}")
    print(f"❌ forecast_engine: {e}")

try:
    from core.nowcast_engine import run_nowcast
    print("✅ nowcast_engine OK")
except Exception as e:
    errors.append(f"nowcast_engine: {e}")
    print(f"❌ nowcast_engine: {e}")

try:
    from core.quarterly import quarterly_mean, calc_ga_trim
    print("✅ quarterly OK")
except Exception as e:
    errors.append(f"quarterly: {e}")
    print(f"❌ quarterly: {e}")

try:
    from core.tempdisagg import chow_lin, denton_cholette
    print("✅ tempdisagg OK")
except Exception as e:
    errors.append(f"tempdisagg: {e}")
    print(f"❌ tempdisagg: {e}")

try:
    from core.cemac_engine import compute_icae_cemac
    print("✅ cemac_engine OK")
except Exception as e:
    errors.append(f"cemac_engine: {e}")
    print(f"❌ cemac_engine: {e}")

try:
    from io_utils.excel_reader import load_country_file
    print("✅ excel_reader OK")
except Exception as e:
    errors.append(f"excel_reader: {e}")
    print(f"❌ excel_reader: {e}")

try:
    from io_utils.excel_writer import write_icae_output
    print("✅ excel_writer OK")
except Exception as e:
    errors.append(f"excel_writer: {e}")
    print(f"❌ excel_writer: {e}")

try:
    from io_utils.word_report import generate_note_icae
    print("✅ word_report OK")
except Exception as e:
    errors.append(f"word_report: {e}")
    print(f"❌ word_report: {e}")

try:
    from ui.charts import chart_icae_monthly, chart_ga_bars
    print("✅ charts OK")
except Exception as e:
    errors.append(f"charts: {e}")
    print(f"❌ charts: {e}")

try:
    from ui.components import download_button
    print("✅ components OK")
except Exception as e:
    errors.append(f"components: {e}")
    print(f"❌ components: {e}")

try:
    from ui.styles import CUSTOM_CSS
    print("✅ styles OK")
except Exception as e:
    errors.append(f"styles: {e}")
    print(f"❌ styles: {e}")

print(f"\n{'='*40}")
if errors:
    print(f"❌ {len(errors)} erreur(s)")
    for e in errors:
        print(f"  - {e}")
else:
    print("✅ Tous les imports sont OK !")

# Test quick data loading
print(f"\n{'='*40}")
print("Test de chargement d'un fichier ICAE...")
from pathlib import Path
test_file = CONSOLIDES / "ICAE_CMR_Consolide.xlsx"
if test_file.exists():
    data = load_country_file(test_file)
    print(f"✅ Fichier CMR chargé : {list(data.keys())}")
    if 'consignes' in data:
        print(f"   Consignes : {data['consignes']}")
    if 'donnees_calcul' in data:
        dc = data['donnees_calcul']
        print(f"   Données calcul : {dc.shape}, colonnes={list(dc.columns[:5])}...")
        print(f"   Période : {dc['Date'].min()} → {dc['Date'].max()}")
    if 'codification' in data:
        cod = data['codification']
        print(f"   Codification : {cod.shape}, colonnes={list(cod.columns)}")
else:
    print(f"⚠️ Fichier non trouvé : {test_file}")
