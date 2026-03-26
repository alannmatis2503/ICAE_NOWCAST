"""Debug: check column names vs codification codes."""
import sys, os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.path.insert(0, r"c:\Users\HP\Documents\Stage pro BEAC\Work\ICAE\Mars 2026\Apps\ICAE_Streamlit")

from config import CONSOLIDES
from io_utils.excel_reader import read_codification, read_donnees_calcul

f = CONSOLIDES / "ICAE_CMR_Consolide.xlsx"
cod = read_codification(f)
don = read_donnees_calcul(f)

print("=== Codification ===")
print(cod[["Code", "Label"]].to_string())
print("\n=== Donnees_calcul columns ===")
for i, c in enumerate(don.columns):
    print(f"  {i}: '{c}'")
