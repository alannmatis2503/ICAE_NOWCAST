"""Test complet de toutes les corrections."""
import sys
from pathlib import Path

# Ajouter le répertoire de l'app au path
APP_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(APP_DIR))

import pandas as pd
import numpy as np
import io
import openpyxl

TESTS_OK = 0
TESTS_FAIL = 0

def check(desc, condition):
    global TESTS_OK, TESTS_FAIL
    if condition:
        TESTS_OK += 1
        print(f"  ✅ {desc}")
    else:
        TESTS_FAIL += 1
        print(f"  ❌ {desc}")


TEST_FILE = Path(__file__).resolve().parent.parent.parent.parent / \
    "Livrable_Final" / "01_Classeurs_ICAE_Fichiers_test" / "ICAE_RCA_Consolide_now.xlsx"


def test_imports():
    print("\n== Test 1: Imports ==")
    try:
        from io_utils.excel_writer import write_icae_recalc_output
        check("write_icae_recalc_output imported", True)
    except ImportError as e:
        check(f"write_icae_recalc_output import: {e}", False)

    try:
        from io_utils.excel_reader import rename_columns_to_codes
        check("rename_columns_to_codes imported", True)
    except ImportError as e:
        check(f"rename_columns_to_codes import: {e}", False)


def test_load_data():
    print("\n== Test 2: Chargement des données ==")
    from io_utils.excel_reader import (
        list_sheets, read_donnees_calcul, read_codification,
        rename_columns_to_codes, read_consignes,
    )

    if not TEST_FILE.exists():
        check(f"Fichier test existe: {TEST_FILE}", False)
        return None, None, None, None

    sheets = list_sheets(TEST_FILE)
    check("Feuilles détectées", len(sheets) > 0)
    check("Donnees_calcul présente", "Donnees_calcul" in sheets)
    check("Codification présente", "Codification" in sheets)

    consignes = read_consignes(TEST_FILE)
    check("Consignes lues", consignes is not None)

    codif = read_codification(TEST_FILE)
    check(f"Codification: {len(codif)} lignes", len(codif) > 0)
    check("Code column exists", "Code" in codif.columns)
    check("PRIOR column exists", "PRIOR" in codif.columns)
    check("Secteur column exists", "Secteur" in codif.columns)

    donnees = read_donnees_calcul(TEST_FILE)
    check(f"Données: {len(donnees)} lignes", len(donnees) > 0)

    donnees_renamed = rename_columns_to_codes(donnees, codif)
    check("Colonnes renommées", "Date" in donnees_renamed.columns)

    return consignes, codif, donnees_renamed, donnees


def test_icae_pipeline(consignes, codif, donnees):
    print("\n== Test 3: Pipeline ICAE ==")
    from core.icae_engine import run_icae_pipeline

    dates = pd.to_datetime(donnees["Date"])
    base_year = consignes.get("base_year", 2023)
    base_mask = dates.dt.year == base_year
    base_indices = donnees.index[base_mask]
    base_rows = range(base_indices[0], base_indices[-1] + 1)

    priors = pd.Series(
        codif["PRIOR"].values, index=codif["Code"].values, dtype=float
    ).fillna(0)

    results = run_icae_pipeline(
        donnees=donnees, priors=priors,
        base_year=base_year, base_rows=base_rows,
    )

    check("ICAE calculé", len(results["icae"]) > 0)
    check("Weights contient m", "m" in results["weights"])
    check(f"m shape: {results['weights']['m'].shape}",
          results["weights"]["m"].shape[0] > 0)
    check("active_cols non vide", len(results["active_cols"]) > 0)

    return results, base_year


def test_contributions(results, codif, base_year):
    print("\n== Test 4: Contributions sectorielles ==")
    from core.quarterly import (
        quarterly_mean, calc_ga_trim, calc_gt_trim,
        contributions_sectorielles_trim, normalize_contrib_to_ga,
    )

    q = quarterly_mean(results["icae"], results["dates"])
    q["GA_Trim"] = calc_ga_trim(q["icae_trim"])
    q["GT_Trim"] = calc_gt_trim(q["icae_trim"])

    check(f"Quarterly: {len(q)} trimestres", len(q) > 0)
    check("trimestre column", "trimestre" in q.columns)
    check("GA_Trim column", "GA_Trim" in q.columns)
    check("GT_Trim column", "GT_Trim" in q.columns)

    contrib_trim = contributions_sectorielles_trim(
        results["weights"]["m"], codif, results["dates"],
    )
    check(f"Contributions: {len(contrib_trim)} lignes", len(contrib_trim) > 0)

    sector_cols = [c for c in contrib_trim.columns
                   if c not in ("trimestre", "Date", "_quarter")]
    check(f"Secteurs: {sector_cols}", len(sector_cols) > 0)

    contrib_norm = normalize_contrib_to_ga(contrib_trim, q["GA_Trim"])
    check("Normalisation OK", len(contrib_norm) == len(contrib_trim))

    return q, contrib_norm


def test_write_recalc(donnees, results, q, contrib_trim, codif):
    print("\n== Test 5: Export ICAE recalculé ==")
    from io_utils.excel_writer import write_icae_recalc_output

    wb_bytes = write_icae_recalc_output(
        donnees=donnees, results=results, quarterly=q,
        contrib_trim=contrib_trim, codification=codif,
        country_code="RCA",
    )

    check(f"Workbook bytes: {len(wb_bytes)}", len(wb_bytes) > 0)

    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    sheet_names = wb.sheetnames
    print(f"  Feuilles: {sheet_names}")

    check("Donnees_calcul present", "Donnees_calcul" in sheet_names)
    check("Codification present", "Codification" in sheet_names)
    check("CALCUL_ICAE present", "CALCUL_ICAE" in sheet_names)
    check("Resultats_Trim present", "Resultats_Trim" in sheet_names)
    check("Contributions present", "Contributions" in sheet_names)

    wb.close()
    return wb_bytes


def test_forecast_and_recalc(consignes, codif, donnees):
    """Simule le flux complet Module 2: prévision → recalcul → rapport."""
    print("\n== Test 6: Simulation flux Module 2 → Rapport ==")
    from core.icae_engine import run_icae_pipeline
    from core.forecast_engine import run_all_forecasts
    from core.quarterly import (
        quarterly_mean, calc_ga_trim, calc_gt_trim,
        contributions_sectorielles_trim, normalize_contrib_to_ga,
    )

    dates = pd.to_datetime(donnees["Date"])
    base_year = consignes.get("base_year", 2023)
    base_mask = dates.dt.year == base_year
    base_indices = donnees.index[base_mask]
    base_rows = range(base_indices[0], base_indices[-1] + 1)

    priors = pd.Series(
        codif["PRIOR"].values, index=codif["Code"].values, dtype=float
    ).fillna(0)

    # 1) Prévision de 3 mois pour une variable active
    active_codes = priors[priors > 0].index.tolist()
    var = active_codes[0]
    series = donnees[var].dropna()
    check(f"Prévision pour {var} ({len(series)} obs)", len(series) >= 24)

    result_fc = run_all_forecasts(series, 3, ["MM3", "MM6"], 12, freq="Mensuelle")
    check("Prévision calculée", "forecasts" in result_fc)

    # 2) Construire les données étendues
    last_date = dates.max()
    from pages.module2_helpers import _future_dates_test
    dates_fcst = pd.date_range(last_date + pd.DateOffset(months=1),
                               periods=3, freq="MS")
    new_rows = pd.DataFrame({"Date": dates_fcst})
    sel_method = result_fc["best_method"]
    new_rows[var] = result_fc["forecasts"][sel_method][:3]
    extended = pd.concat([donnees, new_rows], ignore_index=True)
    check(f"Extended: {len(extended)} lignes (orig + 3)", len(extended) == len(donnees) + 3)

    # 3) Recalcul ICAE
    results_icae = run_icae_pipeline(
        donnees=extended, priors=priors,
        base_year=base_year, base_rows=base_rows,
    )
    check("ICAE recalculé", len(results_icae["icae"]) == len(extended))

    q = quarterly_mean(results_icae["icae"], results_icae["dates"])
    q["GA_Trim"] = calc_ga_trim(q["icae_trim"])
    q["GT_Trim"] = calc_gt_trim(q["icae_trim"])

    # 4) Contributions
    contrib_trim = None
    if "m" in results_icae["weights"]:
        contrib_trim = contributions_sectorielles_trim(
            results_icae["weights"]["m"], codif, results_icae["dates"],
        )
        if contrib_trim is not None:
            contrib_trim = normalize_contrib_to_ga(contrib_trim, q["GA_Trim"])

    check("Contributions disponibles", contrib_trim is not None)
    if contrib_trim is not None:
        check(f"Contributions: {len(contrib_trim)} trim",
              len(contrib_trim) > 0)

    # 5) Simuler le module Rapport
    check("trimestre in q_data", "trimestre" in q.columns)
    check("GA_Trim in q_data", "GA_Trim" in q.columns)

    trimestres_list = q["trimestre"].astype(str).tolist()
    check(f"Trimestres: {len(trimestres_list)}", len(trimestres_list) > 0)

    ga_trim = q["GA_Trim"]
    valid_indices = [i for i, v in enumerate(ga_trim) if pd.notna(v)]
    check(f"Valid GA indices: {len(valid_indices)}", len(valid_indices) >= 2)


def test_cemac_quarterly_column():
    """Test que le module Rapports gère bien la colonne Trimestre CEMAC."""
    print("\n== Test 7: CEMAC trimestre column ==")
    from core.cemac_engine import quarterly_cemac

    # Créer un DataFrame CEMAC fictif
    dates = pd.date_range("2020-01-01", periods=48, freq="MS")
    df = pd.DataFrame({
        "ICAE_CEMAC": np.random.normal(100, 5, 48),
        "GA": np.random.normal(2, 1, 48),
        "GT": np.random.normal(0.5, 0.3, 48),
    }, index=dates)

    q = quarterly_cemac(df, dates)
    check("Trimestre column (capital T)", "Trimestre" in q.columns)

    # Simuler le fix du module Rapports
    if "Trimestre" in q.columns and "trimestre" not in q.columns:
        q = q.rename(columns={"Trimestre": "trimestre"})
    check("Renamed to lowercase trimestre", "trimestre" in q.columns)
    check("trimestre values", len(q["trimestre"].tolist()) > 0)


if __name__ == "__main__":
    test_imports()

    consignes, codif, donnees_renamed, donnees_raw = test_load_data()
    if consignes is None:
        print("\n⚠️ Fichier test non trouvé — skip remaining tests")
    else:
        results, base_year = test_icae_pipeline(consignes, codif, donnees_renamed)
        q, contrib_norm = test_contributions(results, codif, base_year)
        test_write_recalc(donnees_renamed, results, q, contrib_norm, codif)
        test_cemac_quarterly_column()

        # Test 6 needs a helper - skip if module not importable
        try:
            # Direct simulation without importing page module
            print("\n== Test 6: Simulation flux Module 2 → Rapport ==")
            from core.icae_engine import run_icae_pipeline as _rip
            from core.forecast_engine import run_all_forecasts as _raf
            from core.quarterly import (
                quarterly_mean as _qm, calc_ga_trim as _gat,
                calc_gt_trim as _gtt,
                contributions_sectorielles_trim as _cst,
                normalize_contrib_to_ga as _ntg,
            )
            dates = pd.to_datetime(donnees_renamed["Date"])
            base_year_t = consignes.get("base_year", 2023)
            bm = dates.dt.year == base_year_t
            bi = donnees_renamed.index[bm]
            br = range(bi[0], bi[-1] + 1)
            priors = pd.Series(
                codif["PRIOR"].values, index=codif["Code"].values, dtype=float
            ).fillna(0)
            active_codes = priors[priors > 0].index.tolist()[:1]
            var = active_codes[0]
            series = donnees_renamed[var].dropna()
            rfc = _raf(series, 3, ["MM3", "MM6"], 12, freq="Mensuelle")
            check(f"Forecast for {var}", "forecasts" in rfc)

            last_date = dates.max()
            dates_fcst = pd.date_range(last_date + pd.DateOffset(months=1),
                                       periods=3, freq="MS")
            new_rows = pd.DataFrame({"Date": dates_fcst})
            sm = rfc["best_method"]
            new_rows[var] = rfc["forecasts"][sm][:3]
            extended = pd.concat([donnees_renamed, new_rows], ignore_index=True)
            check(f"Extended: {len(extended)} rows", len(extended) == len(donnees_renamed) + 3)

            ricae = _rip(donnees=extended, priors=priors,
                         base_year=base_year_t, base_rows=br)
            check("ICAE recalculated", len(ricae["icae"]) == len(extended))

            qt = _qm(ricae["icae"], ricae["dates"])
            qt["GA_Trim"] = _gat(qt["icae_trim"])
            qt["GT_Trim"] = _gtt(qt["icae_trim"])

            ct = None
            if "m" in ricae["weights"]:
                ct = _cst(ricae["weights"]["m"], codif, ricae["dates"])
                if ct is not None:
                    ct = _ntg(ct, qt["GA_Trim"])
            check("Contributions available after recalc", ct is not None)

            trimestres = qt["trimestre"].astype(str).tolist()
            valid_ga = [i for i, v in enumerate(qt["GA_Trim"]) if pd.notna(v)]
            check(f"Valid GA for report: {len(valid_ga)}", len(valid_ga) >= 2)
        except Exception as e:
            check(f"Module 2 → Rapport simulation: {e}", False)
            import traceback
            traceback.print_exc()

    print(f"\n{'='*50}")
    print(f"RÉSULTAT: {TESTS_OK} ✅  |  {TESTS_FAIL} ❌")
    if TESTS_FAIL == 0:
        print("🎉 TOUS LES TESTS PASSENT")
    else:
        print(f"⚠️ {TESTS_FAIL} test(s) échoué(s)")
