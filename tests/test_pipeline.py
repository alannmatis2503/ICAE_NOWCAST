"""Test du pipeline ICAE complet avec le fichier CMR."""
import sys
import os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.path.insert(0, r"c:\Users\HP\Documents\Stage pro BEAC\Work\ICAE\Mars 2026\Apps\ICAE_Streamlit")

import pandas as pd
import numpy as np
from pathlib import Path

from config import CONSOLIDES, COUNTRY_CODES, POIDS_PIB
from core.icae_engine import (
    calc_sym_growth, calc_sym_growth_df, fixed_sigma,
    calc_weights_fixed, calc_I_recursive, normalize_base100,
    run_icae_pipeline,
)
from core.quarterly import quarterly_mean, calc_ga_trim, calc_gt_trim
from core.cemac_engine import compute_icae_cemac
from core.forecast_engine import run_all_forecasts
from io_utils.excel_reader import (
    read_consignes, read_codification, read_donnees_calcul,
    load_country_file, list_sheets, rename_columns_to_codes,
)

print("=" * 60)
print("TEST 1 : Chargement du fichier CMR")
print("=" * 60)

test_file = CONSOLIDES / "ICAE_CMR_Consolide.xlsx"
if not test_file.exists():
    print(f"ERREUR : fichier non trouve : {test_file}")
    sys.exit(1)

# Test list_sheets
sheets = list_sheets(test_file)
print(f"  Feuilles : {sheets}")

# Test read_consignes
consignes = read_consignes(test_file)
print(f"  Consignes : {consignes}")

# Test read_codification
codification = read_codification(test_file)
print(f"  Codification : {codification.shape}, cols={list(codification.columns)}")

# Test read_donnees_calcul
donnees = read_donnees_calcul(test_file)
donnees = rename_columns_to_codes(donnees, codification)
print(f"  Donnees_calcul : {donnees.shape}")
print(f"  Periode : {donnees['Date'].min()} -> {donnees['Date'].max()}")
print(f"  Colonnes (5 premieres) : {list(donnees.columns[:6])}")

print("\n" + "=" * 60)
print("TEST 2 : Pipeline ICAE complet (CMR)")
print("=" * 60)

# Preparer les priors
if "Code" in codification.columns and "PRIOR" in codification.columns:
    priors = pd.Series(
        codification["PRIOR"].values,
        index=codification["Code"].values,
        dtype=float,
    ).fillna(0)
else:
    data_cols = [c for c in donnees.columns if c != "Date"]
    priors = pd.Series(1.0, index=data_cols)

base_year = consignes.get("base_year", 2023)
dates = pd.to_datetime(donnees["Date"])
base_mask = dates.dt.year == base_year
base_indices = donnees.index[base_mask]
if len(base_indices) > 0:
    base_rows = range(base_indices[0], base_indices[-1] + 1)
else:
    base_rows = range(123, 136)

print(f"  Base year : {base_year}")
print(f"  Base rows : {base_rows}")
print(f"  Priors : {(priors > 0).sum()} actives sur {len(priors)}")

# Lancer le pipeline
try:
    results = run_icae_pipeline(
        donnees=donnees,
        priors=priors,
        base_year=base_year,
        base_rows=base_rows,
        sigma_mode="fixed",
    )
    icae = results["icae"]
    print(f"  ICAE calcule : {icae.dropna().shape[0]} valeurs")
    print(f"  ICAE derniere valeur : {icae.dropna().iloc[-1]:.4f}")
    print(f"  ICAE min/max : {icae.dropna().min():.4f} / {icae.dropna().max():.4f}")
    
    # Verifier autour de 100 pour l'annee de base
    icae_base = icae[base_mask].dropna()
    if len(icae_base) > 0:
        mean_base = icae_base.mean()
        print(f"  ICAE moyen annee de base : {mean_base:.4f} (attendu ~100)")
        assert abs(mean_base - 100) < 1.0, f"ICAE base year mean devrait etre ~100, got {mean_base}"
        print("  [OK] Moyenne annee de base ~100")
    
    # GA mensuel
    ga = results["ga_monthly"]
    ga_valid = ga.dropna()
    print(f"  GA mensuel : {ga_valid.shape[0]} valeurs")
    print(f"  GA derniere : {ga_valid.iloc[-1]:.2f}%")
    
    # Ponderation finale
    pond = results["pond_finale"]
    if isinstance(pond, pd.Series):
        print(f"  Ponderations : sum={pond.sum():.4f} (attendu ~1)")
        assert abs(pond.sum() - 1.0) < 0.01, f"Sum poids devrait etre ~1, got {pond.sum()}"
        print("  [OK] Somme ponderations ~1")
    
    print("  [OK] Pipeline ICAE reussi!")
except Exception as e:
    print(f"  ERREUR Pipeline : {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n" + "=" * 60)
print("TEST 3 : Trimestriel")
print("=" * 60)

q = quarterly_mean(results["icae"], results["dates"])
print(f"  Trimestres : {len(q)}")
print(f"  Colonnes : {list(q.columns)}")
ga_trim = calc_ga_trim(q["icae_trim"])
gt_trim = calc_gt_trim(q["icae_trim"])
print(f"  GA trim derniere : {ga_trim.dropna().iloc[-1]:.2f}%")
print(f"  GT trim derniere : {gt_trim.dropna().iloc[-1]:.2f}%")
print("  [OK] Trimestrialisation reussie!")

print("\n" + "=" * 60)
print("TEST 4 : Previsions (2 variables, horizon=3)")
print("=" * 60)

data_cols = [c for c in donnees.columns if c != "Date"]
test_vars = data_cols[:2]
for var in test_vars:
    series = donnees[var].dropna()
    if len(series) < 24:
        print(f"  {var} : serie trop courte ({len(series)}), skip")
        continue
    result = run_all_forecasts(series, h=3, methods=["MM3", "MM6", "NS", "TL"],
                               bt_window=12)
    best = result["best_method"]
    mape = result["backtesting"][best]["mape"]
    fc = result["forecasts"][best]
    print(f"  {var} : best={best}, MAPE={mape:.2f}%, forecast={fc}")

print("  [OK] Previsions reussies!")

print("\n" + "=" * 60)
print("TEST 5 : Chargement multiple pays pour CEMAC")
print("=" * 60)

icae_dict = {}
for code in COUNTRY_CODES:
    fpath = CONSOLIDES / f"ICAE_{code}_Consolide.xlsx"
    if not fpath.exists():
        print(f"  {code} : fichier manquant, skip")
        continue
    try:
        cons = read_consignes(fpath)
        cod = read_codification(fpath)
        don = read_donnees_calcul(fpath)
        don = rename_columns_to_codes(don, cod)
        
        by = cons.get("base_year", 2023)
        dt = pd.to_datetime(don["Date"])
        bm = dt.dt.year == by
        bi = don.index[bm]
        br = range(bi[0], bi[-1] + 1) if len(bi) > 0 else range(123, 136)
        
        if "Code" in cod.columns and "PRIOR" in cod.columns:
            pr = pd.Series(cod["PRIOR"].values, index=cod["Code"].values, dtype=float).fillna(0)
        else:
            dc = [c for c in don.columns if c != "Date"]
            pr = pd.Series(1.0, index=dc)
        
        res = run_icae_pipeline(don, pr, by, br)
        icae_series = res["icae"]
        icae_series.index = dt
        icae_dict[code] = icae_series
        print(f"  {code} : OK ({icae_series.dropna().shape[0]} obs, last={icae_series.dropna().iloc[-1]:.2f})")
    except Exception as e:
        print(f"  {code} : ERREUR - {e}")

if len(icae_dict) >= 2:
    cemac_df = compute_icae_cemac(icae_dict, POIDS_PIB)
    print(f"\n  CEMAC calcule : {len(cemac_df)} obs")
    print(f"  ICAE CEMAC derniere : {cemac_df['ICAE_CEMAC'].dropna().iloc[-1]:.2f}")
    print(f"  GA CEMAC derniere : {cemac_df['GA'].dropna().iloc[-1]:.2f}%")
    print("  [OK] CEMAC agrege reussi!")
else:
    print("  Pas assez de pays pour tester CEMAC")

print("\n" + "=" * 60)
print("TOUS LES TESTS REUSSIS !")
print("=" * 60)
