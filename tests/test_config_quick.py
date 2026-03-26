"""Test rapide d'importation — modules légers uniquement."""
import sys
import os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

mods = [
    ("config", "from config import COUNTRY_CODES, POIDS_PIB, CONSOLIDES, COUNTRY_NAMES, PIB_2014"),
    ("quarterly", "from core.quarterly import quarterly_mean, calc_ga_trim"),
    ("excel_reader", "from io_utils.excel_reader import load_country_file, read_codification, list_sheets"),
]

errors = []
for name, imp in mods:
    try:
        exec(imp)
        print(f"OK {name}")
    except Exception as e:
        errors.append(f"{name}: {e}")
        print(f"ERREUR {name}: {e}")

# Vérifier les valeurs config
from config import COUNTRY_CODES, PIB_2014, POIDS_PIB

print(f"\nCOUNTRY_CODES = {COUNTRY_CODES}")
print(f"PIB_2014 = {PIB_2014}")
total = sum(PIB_2014.values())
print(f"Total PIB_2014 = {total:.2f}")
for c, v in POIDS_PIB.items():
    print(f"  Poids {c} = {v:.4f}")
print(f"Somme poids = {sum(POIDS_PIB.values()):.6f}")

assert COUNTRY_CODES == ["CMR", "RCA", "CNG", "GAB", "GNQ", "TCD"], "Ordre BEAC incorrect"
assert abs(sum(POIDS_PIB.values()) - 1.0) < 0.01, "Somme poids != 1"
assert abs(PIB_2014["CMR"] - 22310.90) < 1, "PIB CMR incorrect"
print("\nTous les tests config OK")

if errors:
    print(f"\n{len(errors)} erreur(s)")
    sys.exit(1)
else:
    print("Tous les imports OK")
