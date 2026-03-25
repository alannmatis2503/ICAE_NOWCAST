"""Écriture de fichiers Excel avec formules et valeurs."""
import io
import copy
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path


# Styles
HIST_FILL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
HIST_FONT = Font(color="1F4E79")
FCST_FILL = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
FCST_FONT = Font(color="E26B0A", bold=True)
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def write_icae_output(source_path, results: dict, country_code: str) -> bytes:
    """
    Crée un fichier _OUT.xlsx à partir du source avec résultats injectés.
    Resultats_Trim contient des formules Excel d'agrégation référençant CALCUL_ICAE.

    Returns : bytes du fichier Excel
    """
    wb = openpyxl.load_workbook(source_path)

    # ── Mettre à jour CALCUL_ICAE ─────────────────────────────────────────
    icae_col_letter = None
    data_start_row = 16  # les données commencent à la ligne 16 dans CALCUL_ICAE

    if "CALCUL_ICAE" in wb.sheetnames:
        ws = wb["CALCUL_ICAE"]
        icae = results.get("icae")

        if icae is not None:
            max_col = ws.max_column
            for col in range(1, max_col + 1):
                h = ws.cell(row=15, column=col).value
                if h and "ICAE" in str(h):
                    icae_col_letter = get_column_letter(col)
                    for i, val in enumerate(icae.values):
                        if not np.isnan(val):
                            r = data_start_row + i
                            ws.cell(row=r, column=col, value=round(val, 6))

    # ── Mettre à jour Resultats_Trim avec formules Excel ──────────────────
    if "Resultats_Trim" in wb.sheetnames and "quarterly" in results:
        ws_rt = wb["Resultats_Trim"]
        q_data = results["quarterly"]
        dates = results.get("dates")

        # Construire la correspondance trimestre → lignes dans CALCUL_ICAE
        if dates is not None and icae_col_letter is not None:
            dates_ts = pd.to_datetime(dates)
            quarter_groups = {}
            for i, d in enumerate(dates_ts):
                qkey = f"{d.year}T{(d.month - 1) // 3 + 1}"
                if qkey not in quarter_groups:
                    quarter_groups[qkey] = []
                quarter_groups[qkey].append(data_start_row + i)

            # En-têtes pour Resultats_Trim
            headers = ["Trimestre", "ICAE_Trim", "GA_Trim", "GT_Trim"]
            for j, h in enumerate(headers):
                cell = ws_rt.cell(row=1, column=j + 1, value=h)
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.border = THIN_BORDER

            # Écrire les formules pour chaque trimestre
            for idx, (_, row) in enumerate(q_data.iterrows()):
                r = idx + 2
                trim_label = row.get("trimestre", str(row.get("quarter", "")))
                ws_rt.cell(row=r, column=1, value=trim_label).border = THIN_BORDER

                rows_in_q = quarter_groups.get(trim_label, [])

                if rows_in_q and icae_col_letter:
                    # ICAE_Trim = AVERAGE(CALCUL_ICAE!col:col) sur les 3 mois
                    refs = ",".join(
                        f"CALCUL_ICAE!{icae_col_letter}{rr}" for rr in rows_in_q
                    )
                    formula_avg = f"=AVERAGE({refs})"
                    ws_rt.cell(row=r, column=2, value=formula_avg).border = THIN_BORDER

                    # GA_Trim = (ICAE_T / ICAE_{T-4} - 1) * 100
                    if idx >= 4:
                        r_prev_year = idx - 4 + 2
                        col_b = get_column_letter(2)
                        formula_ga = f"=({col_b}{r}/{col_b}{r_prev_year}-1)*100"
                        ws_rt.cell(row=r, column=3, value=formula_ga).border = THIN_BORDER
                    else:
                        ws_rt.cell(row=r, column=3).border = THIN_BORDER

                    # GT_Trim = (ICAE_T / ICAE_{T-1} - 1) * 100
                    if idx >= 1:
                        r_prev = idx - 1 + 2
                        formula_gt = f"=({col_b}{r}/{col_b}{r_prev}-1)*100"
                        ws_rt.cell(row=r, column=4, value=formula_gt).border = THIN_BORDER
                    else:
                        ws_rt.cell(row=r, column=4).border = THIN_BORDER
                else:
                    # Pas de correspondance, écrire les valeurs directement
                    for j, col_name in enumerate(["icae_trim", "GA_Trim", "GT_Trim"]):
                        val = row.get(col_name)
                        if val is not None and not (isinstance(val, float) and np.isnan(val)):
                            ws_rt.cell(row=r, column=j + 2, value=round(float(val), 4))
                        ws_rt.cell(row=r, column=j + 2).border = THIN_BORDER

    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    buf.seek(0)
    return buf.getvalue()


def write_previsions_excel(donnees: pd.DataFrame, previsions: dict,
                           backtesting: dict, recommandations: dict,
                           scenarios: dict = None,
                           edited: pd.DataFrame = None) -> bytes:
    """Génère le fichier Excel de prévisions avec séries complètes."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        # Feuille Donnees_Completes : historique + prévisions (meilleure méthode)
        # pour que l'utilisateur ait les séries complètes en une seule feuille
        if edited is not None:
            complete = donnees.copy()
            for col in edited.columns:
                if col != "Date" and col in complete.columns:
                    # Les prévisions éditées ont des dates formatées en string ;
                    # on les ajoute à la fin du DataFrame
                    pass
            # Utiliser le DataFrame édité directement
            import pandas as _pd
            last_date = _pd.to_datetime(complete["Date"]).max()
            n_fc = len(edited)
            # Construire les dates futures selon la fréquence
            fc_dates = _pd.date_range(last_date + _pd.DateOffset(months=1),
                                      periods=n_fc, freq="MS")
            new_rows = _pd.DataFrame({"Date": fc_dates})
            for col in edited.columns:
                if col != "Date" and col in complete.columns:
                    new_rows[col] = edited[col].values
            extended = _pd.concat([complete, new_rows], ignore_index=True)
            extended.to_excel(writer, sheet_name="Donnees_Completes", index=False)
        else:
            donnees.to_excel(writer, sheet_name="Donnees_Completes", index=False)

        # Feuille Donnees (historique seul)
        donnees.to_excel(writer, sheet_name="Donnees", index=False)

        # Feuille Previsions
        prev_rows = []
        for var, methods in previsions.items():
            for method, values in methods.items():
                for i, v in enumerate(values):
                    prev_rows.append({
                        "Variable": var, "Methode": method,
                        "Mois": i + 1, "Valeur": v,
                    })
        if prev_rows:
            pd.DataFrame(prev_rows).to_excel(
                writer, sheet_name="Previsions", index=False)

        # Feuille Backtesting
        bt_rows = []
        for var, methods in backtesting.items():
            row = {"Variable": var}
            for method, metrics in methods.items():
                row[f"{method}_MAPE"] = metrics.get("mape", np.nan)
                row[f"{method}_MAE"] = metrics.get("mae", np.nan)
                row[f"{method}_RMSE"] = metrics.get("rmse", np.nan)
            bt_rows.append(row)
        if bt_rows:
            pd.DataFrame(bt_rows).to_excel(
                writer, sheet_name="Backtesting", index=False)

        # Feuille Recommandations
        rec_rows = []
        for var, info in recommandations.items():
            rec_rows.append({
                "Variable": var,
                "Methode_recommandee": info["method"],
                "MAPE": info.get("mape", np.nan),
            })
        if rec_rows:
            pd.DataFrame(rec_rows).to_excel(
                writer, sheet_name="Recommandations", index=False)

        # Feuille Scenarios
        if scenarios is not None:
            scenarios.to_excel(writer, sheet_name="ICAE_Scenarios", index=False)

        # Feuille Previsions_Editees
        if edited is not None:
            edited.to_excel(writer, sheet_name="Previsions_Editees",
                            index=False)

    buf.seek(0)
    return buf.getvalue()


def write_nowcast_excel(pib_q: pd.Series, results: dict,
                        params: dict = None) -> bytes:
    """Génère le fichier Excel de résultats Nowcast."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        # PIB et Nowcasts
        df = pd.DataFrame({"PIB_observe": pib_q})
        for name, r in results.items():
            fc = r["forecast"]
            df[name] = fc.reindex(df.index)
        df.to_excel(writer, sheet_name="PIB_and_Nowcasts")

        # Performance
        perf_rows = []
        for name, r in results.items():
            m = r["metrics"]
            perf_rows.append({
                "Modele": name,
                "RMSE_in": m["in_sample"].get("rmse", np.nan),
                "MAE_in": m["in_sample"].get("mae", np.nan),
                "MAPE_in": m["in_sample"].get("mape", np.nan),
                "RMSE_out": m["out_sample"].get("rmse", np.nan),
                "MAE_out": m["out_sample"].get("mae", np.nan),
                "MAPE_out": m["out_sample"].get("mape", np.nan),
                "Correlation": r.get("correlation", np.nan),
            })
        pd.DataFrame(perf_rows).to_excel(
            writer, sheet_name="Performance", index=False)

        # GA
        if len(df) > 4:
            ga_df = pd.DataFrame(index=df.index)
            for col in df.columns:
                ga_df[f"GA_{col}"] = (df[col] / df[col].shift(4) - 1) * 100
            ga_df.to_excel(writer, sheet_name="Glissement_Annuel")

        # Paramètres
        if params:
            pd.DataFrame([params]).to_excel(
                writer, sheet_name="Parametres", index=False)

    buf.seek(0)
    return buf.getvalue()


def _copy_cell_style(src_cell, dst_cell):
    """Copie le style (fill, font, border, alignment, number_format) d'une cellule vers une autre."""
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def _detect_calcul_layout(ws):
    """Détecte la disposition des colonnes dans CALCUL_ICAE.
    Returns dict avec les indices (1-indexed) des colonnes spéciales."""
    layout = {"var_first": 2, "var_last": 2}
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        h14 = str(ws.cell(14, col).value or "")
        h15 = str(ws.cell(15, col).value or "")
        if "Σm" in h14 or "Σm" in h15 or "sum_m" in h14.lower():
            layout["sum_m"] = col
        elif "Indice" in h14 or "Indice" in h15:
            layout["indice"] = col
        elif "Base" in h14 or "Base" in h15:
            layout["base"] = col
        elif "ICAE" in h14 or "ICAE" in h15:
            layout["icae"] = col
        elif "GA" in h14 or "GA" in h15:
            layout["ga"] = col
        elif "GT" in h14 or "GT" in h15:
            layout["gt"] = col
    # Déterminer la plage de colonnes variables (entre col 2 et sum_m - 1)
    if "sum_m" in layout:
        # Colonnes de variables = B jusqu'à la colonne avant un éventuel gap
        # On cherche la dernière colonne avant sum_m qui a un en-tête en ligne 14
        for c in range(layout["sum_m"] - 1, 1, -1):
            v14 = ws.cell(14, c).value
            if v14 is not None and str(v14).strip():
                layout["var_last"] = c
                break
    return layout


def write_icae_recalc_output(donnees: pd.DataFrame,
                             results: dict,
                             quarterly: pd.DataFrame,
                             contrib_trim: pd.DataFrame | None,
                             codification: pd.DataFrame | None,
                             country_code: str,
                             source_path=None) -> bytes:
    """
    Génère un classeur ICAE recalculé (après réinjection de prévisions).

    Si *source_path* est fourni (fichier original), le classeur est copié
    en préservant **toute la mise en forme** et les **formules** existantes ;
    les feuilles Donnees_calcul et CALCUL_ICAE sont allongées avec les
    nouvelles lignes (prévisions).  Les feuilles Resultats_Trim et Contrib
    sont mises à jour.

    Si *source_path* est None, un classeur est créé de zéro (fallback).
    """
    if source_path is not None:
        return _write_icae_from_template(
            source_path, donnees, results, quarterly,
            contrib_trim, codification, country_code,
        )
    return _write_icae_from_scratch(
        donnees, results, quarterly,
        contrib_trim, codification, country_code,
    )


def _write_icae_from_template(source_path,
                               donnees: pd.DataFrame,
                               results: dict,
                               quarterly: pd.DataFrame,
                               contrib_trim: pd.DataFrame | None,
                               codification: pd.DataFrame | None,
                               country_code: str) -> bytes:
    """Génère le classeur en partant du fichier original (mise en forme + formules)."""
    if hasattr(source_path, "seek"):
        source_path.seek(0)
    wb = openpyxl.load_workbook(source_path)

    n_new = 0  # nombre de lignes de prévision à ajouter (calculé plus bas)

    # ── 1) Étendre Donnees_calcul ─────────────────────────────────────────
    if "Donnees_calcul" in wb.sheetnames:
        ws_d = wb["Donnees_calcul"]
        orig_max = ws_d.max_row            # ex: 184 (1 header + 183 data)
        new_total = len(donnees) + 1       # header + données étendues

        if new_total > orig_max:
            # La ligne de référence dont on copie le style
            ref_row = orig_max
            for r_idx in range(orig_max + 1, new_total + 1):
                data_idx = r_idx - 2       # index dans le DataFrame (0-based)
                if data_idx >= len(donnees):
                    break
                for c_idx in range(1, len(donnees.columns) + 1):
                    val = donnees.iloc[data_idx, c_idx - 1]
                    cell = ws_d.cell(row=r_idx, column=c_idx)
                    # Écrire la valeur
                    if pd.isna(val):
                        cell.value = None
                    elif c_idx == 1:
                        cell.value = pd.to_datetime(val)
                        cell.number_format = "YYYY-MM-DD"
                    else:
                        cell.value = float(val) if not pd.isna(val) else None
                    # Copier le style de la ligne de référence
                    _copy_cell_style(ws_d.cell(ref_row, c_idx), cell)
                    # Surligner les lignes de prévision
                    cell.fill = FCST_FILL
                    cell.font = FCST_FONT

    # ── 2) Étendre CALCUL_ICAE ────────────────────────────────────────────
    if "CALCUL_ICAE" in wb.sheetnames:
        ws_c = wb["CALCUL_ICAE"]
        layout = _detect_calcul_layout(ws_c)

        # Déterminer le nombre de lignes originales vs étendues dans TCS
        data_start_calc = 16
        orig_calc_max = ws_c.max_row       # dernière ligne TCS existante
        orig_n_data = orig_max - 1 if "Donnees_calcul" in wb.sheetnames else 183
        ext_n_data = len(donnees)
        n_new = ext_n_data - orig_n_data   # lignes à ajouter

        if n_new > 0:
            var_last = layout.get("var_last", 22)

            for offset in range(n_new):
                calc_row = orig_calc_max + 1 + offset
                d_row_curr = orig_n_data + 1 + offset + 1  # +1 pour header Excel
                d_row_prev = d_row_curr - 1

                # Ligne de référence pour le style
                ref_calc_row = orig_calc_max

                # Col A : Date → formule
                cell_a = ws_c.cell(calc_row, 1)
                cell_a.value = f"=Donnees_calcul!A{d_row_curr}"
                _copy_cell_style(ws_c.cell(ref_calc_row, 1), cell_a)
                cell_a.fill = FCST_FILL
                cell_a.font = FCST_FONT

                # Cols B..var_last : formules TCS
                for vc in range(2, var_last + 1):
                    cl = get_column_letter(vc)
                    cell = ws_c.cell(calc_row, vc)
                    cell.value = (
                        f'=IFERROR(2*(Donnees_calcul!{cl}{d_row_curr}'
                        f'-Donnees_calcul!{cl}{d_row_prev})'
                        f'/(Donnees_calcul!{cl}{d_row_curr}'
                        f'+Donnees_calcul!{cl}{d_row_prev}),"")'
                    )
                    _copy_cell_style(ws_c.cell(ref_calc_row, vc), cell)
                    cell.fill = FCST_FILL

                # Σm
                if "sum_m" in layout:
                    sc = layout["sum_m"]
                    fl = get_column_letter(2)
                    ll = get_column_letter(var_last)
                    cell = ws_c.cell(calc_row, sc)
                    cell.value = (
                        f"=IFERROR(SUMPRODUCT({fl}{calc_row}:{ll}{calc_row},"
                        f"${fl}$12:${ll}$12),\"\")"
                    )
                    _copy_cell_style(ws_c.cell(ref_calc_row, sc), cell)
                    cell.fill = FCST_FILL

                # Indice
                if "indice" in layout and "sum_m" in layout:
                    ic = layout["indice"]
                    scl = get_column_letter(layout["sum_m"])
                    icl = get_column_letter(ic)
                    cell = ws_c.cell(calc_row, ic)
                    cell.value = (
                        f"=IFERROR({icl}{calc_row - 1}"
                        f"*(2+{scl}{calc_row})/(2-{scl}{calc_row}),\"\")"
                    )
                    _copy_cell_style(ws_c.cell(ref_calc_row, ic), cell)
                    cell.fill = FCST_FILL

                # Base (même formule que les autres lignes — moyenne fixe)
                if "base" in layout and "indice" in layout:
                    bc = layout["base"]
                    # Lire la formule Base existante (elle est la même pour toutes les lignes)
                    base_formula = ws_c.cell(ref_calc_row, bc).value
                    cell = ws_c.cell(calc_row, bc)
                    if base_formula and str(base_formula).startswith("="):
                        cell.value = base_formula
                    _copy_cell_style(ws_c.cell(ref_calc_row, bc), cell)
                    cell.fill = FCST_FILL

                # ICAE
                if "icae" in layout and "indice" in layout and "base" in layout:
                    xc = layout["icae"]
                    xcl = get_column_letter(layout["indice"])
                    bcl = get_column_letter(layout["base"])
                    cell = ws_c.cell(calc_row, xc)
                    cell.value = f"=100*{xcl}{calc_row}/{bcl}{calc_row}"
                    _copy_cell_style(ws_c.cell(ref_calc_row, xc), cell)
                    cell.fill = FCST_FILL

                # GA (glissement annuel mensuel : 12 lignes avant)
                if "ga" in layout and "icae" in layout:
                    gc = layout["ga"]
                    icl2 = get_column_letter(layout["icae"])
                    cell = ws_c.cell(calc_row, gc)
                    if calc_row - 12 >= data_start_calc:
                        cell.value = (
                            f'=IFERROR(({icl2}{calc_row}'
                            f'/{icl2}{calc_row - 12}-1)*100,"")'
                        )
                    else:
                        cell.value = '=""'
                    _copy_cell_style(ws_c.cell(ref_calc_row, gc), cell)
                    cell.fill = FCST_FILL

                # GT (glissement mensuel)
                if "gt" in layout:
                    tc = layout["gt"]
                    icl3 = get_column_letter(layout.get("icae", tc - 2))
                    cell = ws_c.cell(calc_row, tc)
                    if calc_row > data_start_calc + 1:
                        cell.value = (
                            f'=IFERROR(({icl3}{calc_row}'
                            f'/{icl3}{calc_row - 1}-1)*100,"")'
                        )
                    else:
                        cell.value = '=""'
                    _copy_cell_style(ws_c.cell(ref_calc_row, tc), cell)
                    cell.fill = FCST_FILL

        # Écrire aussi les VALEURS Python dans la colonne ICAE (double sécurité)
        if "icae" in layout:
            icae_series = results.get("icae")
            if icae_series is not None:
                xc = layout["icae"]
                for i, val in enumerate(icae_series.values):
                    r = data_start_calc + i
                    # Ne PAS écraser les formules des nouvelles lignes
                    # mais écrire dans les lignes existantes qui avaient
                    # déjà des valeurs (pas des formules)
                    existing = ws_c.cell(r, xc).value
                    if existing is None or (
                        isinstance(existing, str) and not existing.startswith("=")
                    ):
                        if not np.isnan(val):
                            ws_c.cell(r, xc, value=round(val, 6))

    # ── 3) Mettre à jour Resultats_Trim ───────────────────────────────────
    if "Resultats_Trim" in wb.sheetnames and quarterly is not None:
        ws_rt = wb["Resultats_Trim"]
        q = quarterly
        # Déterminer la colonne ICAE dans CALCUL_ICAE
        icae_col_letter = None
        if "CALCUL_ICAE" in wb.sheetnames:
            ws_ck = wb["CALCUL_ICAE"]
            ly = _detect_calcul_layout(ws_ck)
            if "icae" in ly:
                icae_col_letter = get_column_letter(ly["icae"])

        dates_all = pd.to_datetime(results.get("dates", donnees["Date"]))
        quarter_groups = {}
        for i, d in enumerate(dates_all):
            qkey = f"{d.year}T{(d.month - 1) // 3 + 1}"
            if qkey not in quarter_groups:
                quarter_groups[qkey] = []
            quarter_groups[qkey].append(data_start_calc + i)

        # Écrire les données trimestrielles (garder la ligne 1 d'en-têtes)
        for idx in range(len(q)):
            r = idx + 2
            row = q.iloc[idx]

            # Trimestre label
            trim_label = str(row.get("trimestre", row.get("quarter", "")))
            cell_t = ws_rt.cell(r, 1)
            cell_t.value = trim_label
            if idx == 0:
                _copy_cell_style(ws_rt.cell(2, 1) if ws_rt.max_row >= 2 else cell_t, cell_t)

            rows_in_q = quarter_groups.get(trim_label, [])

            if rows_in_q and icae_col_letter:
                # ICAE trimestriel = AVERAGE des 3 mois
                refs = ",".join(f"CALCUL_ICAE!{icae_col_letter}{rr}" for rr in rows_in_q)
                ws_rt.cell(r, 2).value = row.get("debut", "")
                ws_rt.cell(r, 3).value = row.get("fin", "")
                ws_rt.cell(r, 4).value = f"=AVERAGE({refs})"

                # GA Trim (T/T-4)
                if idx >= 4:
                    r_prev_y = idx - 4 + 2
                    ws_rt.cell(r, 5).value = f"=(D{r}/D{r_prev_y}-1)*100"
                else:
                    ws_rt.cell(r, 5).value = '=""'

                # GT Trim (T/T-1)
                if idx >= 1:
                    r_prev = idx - 1 + 2
                    ws_rt.cell(r, 6).value = f"=(D{r}/D{r_prev}-1)*100"
                else:
                    ws_rt.cell(r, 6).value = '=""'
            else:
                # Fallback : écrire les valeurs directement
                for j, col_name in enumerate(["debut", "fin", "icae_trim", "GA_Trim", "GT_Trim"]):
                    val = row.get(col_name)
                    if val is not None and not (isinstance(val, float) and np.isnan(val)):
                        ws_rt.cell(r, j + 2).value = val

            # Copier le style de la première ligne de données
            for c in range(1, ws_rt.max_column + 1):
                _copy_cell_style(ws_rt.cell(2, c), ws_rt.cell(r, c))

    # ── 4) Étendre Contrib (formules mensuelles) ─────────────────────────
    if "Contrib" in wb.sheetnames and n_new > 0:
        ws_ct = wb["Contrib"]
        contrib_orig_max = ws_ct.max_row
        ref_contrib_row = contrib_orig_max  # dernière ligne existante

        # Détecter l'offset calc_row = contrib_row + offset
        # en utilisant la ligne 2 (toujours fiable avec des formules complètes)
        import re as _re
        _offset = 15  # valeur par défaut
        for _probe_row in (2, 3, ref_contrib_row):
            _sample = str(ws_ct.cell(_probe_row, 1).value or "")
            _m = _re.search(r"CALCUL_ICAE!A(\d+)", _sample)
            if _m:
                _offset = int(_m.group(1)) - _probe_row
                break

        # Analyser la ligne 2 pour déterminer les types de colonnes :
        #  - CALCUL_ICAE! → contribution individuelle
        #  - formule sans CALCUL_ICAE! → agrégat sectoriel
        #  - None → séparatrice
        var_last_ct = ws_ct.max_column
        _col_types = {}  # col_idx -> "contrib" | "sector" | "sep"
        _sector_formulas = {}  # col_idx -> template formula avec {ROW}

        # Utiliser la ligne 2 (la plus fiable : toutes les formules sont complètes)
        for ci in range(2, var_last_ct + 1):
            fval_r2 = ws_ct.cell(2, ci).value
            if fval_r2 is None:
                _col_types[ci] = "sep"
            elif isinstance(fval_r2, str) and "CALCUL_ICAE!" in fval_r2:
                _col_types[ci] = "contrib"
            elif isinstance(fval_r2, str) and fval_r2.startswith("="):
                _col_types[ci] = "sector"
                # Paramétrer la formule en remplaçant les refs de ligne 2 par {ROW}
                # Ex: =B2+C2+D2 → =B{ROW}+C{ROW}+D{ROW}
                _sector_formulas[ci] = _re.sub(
                    r"(?<=[A-Z])2(?=\D|$)", "{ROW}", fval_r2
                )
            else:
                _col_types[ci] = "sep"

        for offset_i in range(n_new):
            ct_row = contrib_orig_max + 1 + offset_i
            calc_row = ct_row + _offset

            # Col A : date depuis CALCUL_ICAE
            cell_a = ws_ct.cell(ct_row, 1)
            cell_a.value = f"=CALCUL_ICAE!A{calc_row}"
            _copy_cell_style(ws_ct.cell(ref_contrib_row, 1), cell_a)
            cell_a.fill = FCST_FILL
            cell_a.font = FCST_FONT

            # Cols B..max
            for ci in range(2, var_last_ct + 1):
                cell = ws_ct.cell(ct_row, ci)
                _copy_cell_style(ws_ct.cell(ref_contrib_row, ci), cell)

                ctype = _col_types.get(ci, "sep")
                if ctype == "contrib":
                    cl = get_column_letter(ci)
                    cell.value = f"=CALCUL_ICAE!{cl}$12*CALCUL_ICAE!{cl}{calc_row}"
                elif ctype == "sector":
                    cell.value = _sector_formulas[ci].replace("{ROW}", str(ct_row))
                else:
                    cell.value = None  # séparatrice

                cell.fill = FCST_FILL
                cell.font = FCST_FONT

    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    buf.seek(0)
    return buf.getvalue()


def _write_icae_from_scratch(donnees: pd.DataFrame,
                              results: dict,
                              quarterly: pd.DataFrame,
                              contrib_trim: pd.DataFrame | None,
                              codification: pd.DataFrame | None,
                              country_code: str) -> bytes:
    """Fallback : génère un classeur de zéro quand aucun fichier source n'est disponible."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": "#4472C4",
            "border": 1, "align": "center",
        })

        donnees.to_excel(writer, sheet_name="Donnees_calcul", index=False)
        ws = writer.sheets["Donnees_calcul"]
        for j, col in enumerate(donnees.columns):
            ws.write(0, j, col, header_fmt)

        if codification is not None and not codification.empty:
            codification.to_excel(writer, sheet_name="Codification", index=False)
            ws_c = writer.sheets["Codification"]
            for j, col in enumerate(codification.columns):
                ws_c.write(0, j, col, header_fmt)

        dates = results.get("dates", donnees["Date"])
        icae = results["icae"]
        indice = results.get("indice", icae)
        sum_m = results.get("sum_m", pd.Series(dtype=float))
        ga = results.get("ga_monthly", pd.Series(dtype=float))
        calc_df = pd.DataFrame({
            "Date": pd.to_datetime(dates).values,
            "Σm": sum_m.values if len(sum_m) == len(dates) else np.full(len(dates), np.nan),
            "Indice": indice.values if len(indice) == len(dates) else np.full(len(dates), np.nan),
            "ICAE": icae.values if len(icae) == len(dates) else np.full(len(dates), np.nan),
            "GA (%)": ga.values if len(ga) == len(dates) else np.full(len(dates), np.nan),
        })
        calc_df.to_excel(writer, sheet_name="CALCUL_ICAE", index=False)
        ws_ic = writer.sheets["CALCUL_ICAE"]
        for j, col in enumerate(calc_df.columns):
            ws_ic.write(0, j, col, header_fmt)

        if quarterly is not None and not quarterly.empty:
            q_export = quarterly.copy()
            for col in ("quarter", "debut", "fin"):
                if col in q_export.columns:
                    q_export[col] = q_export[col].astype(str)
            q_export.to_excel(writer, sheet_name="Resultats_Trim", index=False)
            ws_rt = writer.sheets["Resultats_Trim"]
            for j, col in enumerate(q_export.columns):
                ws_rt.write(0, j, col, header_fmt)

        if contrib_trim is not None and not contrib_trim.empty:
            contrib_trim.to_excel(writer, sheet_name="Contributions", index=False)
            ws_ct = writer.sheets["Contributions"]
            for j, col in enumerate(contrib_trim.columns):
                ws_ct.write(0, j, col, header_fmt)

    buf.seek(0)
    return buf.getvalue()


def write_cemac_excel(df_monthly: pd.DataFrame,
                      df_quarterly: pd.DataFrame,
                      poids: dict,
                      template_path=None) -> bytes:
    """
    Génère le fichier Excel CEMAC.

    Si *template_path* est fourni, le classeur original est chargé et mis à
    jour tout en conservant la mise en forme et les formules.
    Sinon, un classeur est créé de zéro (fallback).
    """
    if template_path is not None and Path(template_path).exists():
        return _write_cemac_from_template(
            template_path, df_monthly, df_quarterly, poids,
        )
    return _write_cemac_from_scratch(df_monthly, df_quarterly, poids)


# ── Ordre des pays dans les colonnes du classeur CEMAC ────────────────────
_CEMAC_COUNTRIES = ["CMR", "RCA", "CNG", "GAB", "GNQ", "TCD"]
_BRUT_COLS = [3, 4, 5, 6, 7, 8]           # C H  (ICAE_Pays brut)
_REBAS_COLS = [9, 10, 11, 12, 13, 14]     # I N
_ICAE_COL = 15                             # O    ICAE CEMAC
_GA_COL = 16                               # P    GA
_GT_COL = 17                               # Q    GT


def _write_cemac_from_template(template_path,
                                df_monthly: pd.DataFrame,
                                df_quarterly: pd.DataFrame,
                                poids: dict) -> bytes:
    """Charge le classeur CEMAC original et met à jour les données."""
    wb = openpyxl.load_workbook(template_path)

    n_months = len(df_monthly)
    dates_idx = pd.to_datetime(df_monthly.index)
    data_start = 5                          # première ligne de données dans ICAE_Pays
    new_max = data_start + n_months - 1     # dernière ligne de données

    # ── 1) Poids_PIB ─────────────────────────────────────────────────────
    if "Poids_PIB" in wb.sheetnames:
        ws_p = wb["Poids_PIB"]
        from config import PIB_2014
        for i, code in enumerate(_CEMAC_COUNTRIES):
            r = 4 + i
            ws_p.cell(r, 3).value = PIB_2014.get(code, 0)
            ws_p.cell(r, 4).value = f"=C{r}/SUM($C$4:$C$9)"

    # ── 2) ICAE_Pays ─────────────────────────────────────────────────────
    if "ICAE_Pays" in wb.sheetnames:
        ws_m = wb["ICAE_Pays"]

        # Ligne de référence pour copier les styles
        orig_max = ws_m.max_row
        ref_row = min(orig_max, data_start + 1)

        for i in range(n_months):
            r = data_start + i
            dt = dates_idx[i]
            row_data = df_monthly.iloc[i]

            # Date + Mois
            ws_m.cell(r, 1).value = dt.to_pydatetime()
            ws_m.cell(r, 1).number_format = "YYYY-MM-DD"
            ws_m.cell(r, 2).value = dt.month

            # Brut (valeurs)
            for ci, code in enumerate(_CEMAC_COUNTRIES):
                col_idx = _BRUT_COLS[ci]
                val = row_data.get(code)
                if val is not None and not (isinstance(val, float) and np.isnan(val)):
                    ws_m.cell(r, col_idx).value = float(val)
                else:
                    ws_m.cell(r, col_idx).value = None

            # Rebasé (formules ajustées à la plage)
            for ci, code in enumerate(_CEMAC_COUNTRIES):
                brut_cl = get_column_letter(_BRUT_COLS[ci])
                ws_m.cell(r, _REBAS_COLS[ci]).value = (
                    f"=IFERROR(100*{brut_cl}{r}"
                    f"/(SUMPRODUCT((${brut_cl}${data_start}:${brut_cl}${new_max})"
                    f"*(YEAR($A${data_start}:$A${new_max})=AnnéeBase!$B$3))"
                    f'/SUMPRODUCT((YEAR($A${data_start}:$A${new_max})=AnnéeBase!$B$3)*1)),"")'
                )

            # ICAE CEMAC (moyenne pondérée)
            parts_num = "+".join(
                f"IF(ISNUMBER({get_column_letter(_REBAS_COLS[ci])}{r}),"
                f"{get_column_letter(_REBAS_COLS[ci])}{r}*Poids_PIB!$C${4+ci},0)"
                for ci in range(6)
            )
            parts_den = "+".join(
                f"IF(ISNUMBER({get_column_letter(_REBAS_COLS[ci])}{r}),"
                f"Poids_PIB!$C${4+ci},0)"
                for ci in range(6)
            )
            ws_m.cell(r, _ICAE_COL).value = (
                f'=IFERROR(({parts_num})/({parts_den}),"")'
            )

            # GA (12 mois)
            if i >= 12:
                ws_m.cell(r, _GA_COL).value = (
                    f'=IFERROR((O{r}/O{r - 12}-1)*100,"")'
                )
            else:
                ws_m.cell(r, _GA_COL).value = '=""'

            # GT (1 mois)
            if i >= 1:
                ws_m.cell(r, _GT_COL).value = (
                    f'=IFERROR((O{r}/O{r - 1}-1)*100,"")'
                )
            else:
                ws_m.cell(r, _GT_COL).value = '=""'

            # Copier styles de la ligne de référence
            for c in range(1, 18):
                _copy_cell_style(ws_m.cell(ref_row, c), ws_m.cell(r, c))

        # Effacer les lignes en surplus si le template en avait plus
        for r in range(new_max + 1, orig_max + 1):
            for c in range(1, 18):
                ws_m.cell(r, c).value = None

    # ── 3) ICAE_Trimestriel ──────────────────────────────────────────────
    if "ICAE_Trimestriel" in wb.sheetnames:
        ws_q = wb["ICAE_Trimestriel"]
        n_q = len(df_quarterly)
        q_start = 4

        orig_q_max = ws_q.max_row
        ref_qrow = min(orig_q_max, q_start + 1)

        for i in range(n_q):
            r = q_start + i
            row = df_quarterly.iloc[i]
            trim_label = str(row.get("Trimestre", ""))

            # Trimestre label
            ws_q.cell(r, 1).value = trim_label

            # Début / Fin
            try:
                period = pd.Period(trim_label)
                ws_q.cell(r, 2).value = period.start_time.to_pydatetime()
                ws_q.cell(r, 2).number_format = "YYYY-MM-DD"
                ws_q.cell(r, 3).value = period.end_time.to_pydatetime()
                ws_q.cell(r, 3).number_format = "YYYY-MM-DD"
            except Exception:
                ws_q.cell(r, 2).value = None
                ws_q.cell(r, 3).value = None

            # Brut AVERAGEIFS (cols 4-9) → ICAE_Pays cols C-H
            for ci in range(6):
                pays_col_letter = get_column_letter(_BRUT_COLS[ci])
                ws_q.cell(r, 4 + ci).value = (
                    f'=IFERROR(AVERAGEIFS(ICAE_Pays!{pays_col_letter}:{pays_col_letter},'
                    f'ICAE_Pays!$A:$A,">="&B{r},'
                    f'ICAE_Pays!$A:$A,"<="&C{r}),"")'
                )

            # Rebasé AVERAGEIFS (cols 10-15) → ICAE_Pays cols I-N
            for ci in range(6):
                rebas_cl = get_column_letter(_REBAS_COLS[ci])
                ws_q.cell(r, 10 + ci).value = (
                    f'=IFERROR(AVERAGEIFS(ICAE_Pays!{rebas_cl}:{rebas_cl},'
                    f'ICAE_Pays!$A:$A,">="&B{r},'
                    f'ICAE_Pays!$A:$A,"<="&C{r}),"")'
                )

            # ICAE CEMAC quarterly (col 16 = P)
            parts_num_q = "+".join(
                f"IF(ISNUMBER({get_column_letter(10+ci)}{r}),"
                f"{get_column_letter(10+ci)}{r}*Poids_PIB!$C${4+ci},0)"
                for ci in range(6)
            )
            parts_den_q = "+".join(
                f"IF(ISNUMBER({get_column_letter(10+ci)}{r}),"
                f"Poids_PIB!$C${4+ci},0)"
                for ci in range(6)
            )
            ws_q.cell(r, 16).value = f'=IFERROR(({parts_num_q})/({parts_den_q}),"")'

            # GA Trim (4 trimestres)
            if i >= 4:
                ws_q.cell(r, 17).value = f'=IFERROR((P{r}/P{r - 4}-1),"")'
            else:
                ws_q.cell(r, 17).value = '=""'

            # GT Trim (1 trimestre)
            if i >= 1:
                ws_q.cell(r, 18).value = f'=IFERROR((P{r}/P{r - 1}-1),"")'
            else:
                ws_q.cell(r, 18).value = '=""'

            for c in range(1, 19):
                _copy_cell_style(ws_q.cell(ref_qrow, c), ws_q.cell(r, c))

        for r in range(q_start + n_q, orig_q_max + 1):
            for c in range(1, 19):
                ws_q.cell(r, c).value = None

    # ── 4) Contributions_Mens ─────────────────────────────────────────────
    if "Contributions_Mens" in wb.sheetnames:
        ws_cm = wb["Contributions_Mens"]
        cm_start = 4
        orig_cm_max = ws_cm.max_row
        ref_cm = min(orig_cm_max, cm_start + 1)

        for i in range(n_months):
            r = cm_start + i
            dt = dates_idx[i]
            ws_cm.cell(r, 1).value = dt.to_pydatetime()
            ws_cm.cell(r, 1).number_format = "YYYY-MM-DD"

            # ICAE rebasé par pays (cols 2-7) = refs vers ICAE_Pays cols I-N
            for ci in range(6):
                rebas_cl = get_column_letter(_REBAS_COLS[ci])
                pays_r = data_start + i
                ws_cm.cell(r, 2 + ci).value = f"=ICAE_Pays!{rebas_cl}{pays_r}"

            # Poids PIB (cols 8-13) — poids dynamiques basés sur disponibilité
            for ci in range(6):
                brut_cl = get_column_letter(2 + ci)
                poids_parts = "+".join(
                    f"IF(ISNUMBER({get_column_letter(2 + cj)}{r}),Poids_PIB!$C${4+cj},0)"
                    for cj in range(6)
                )
                ws_cm.cell(r, 8 + ci).value = (
                    f"=IFERROR(IF(ISNUMBER({brut_cl}{r}),"
                    f"Poids_PIB!$C${4+ci}/({poids_parts}),0),0)"
                )

            # Contrib GT mensuel (cols 15-20)
            for ci in range(6):
                bcl = get_column_letter(2 + ci)
                pcl = get_column_letter(8 + ci)
                if i >= 1:
                    ws_cm.cell(r, 15 + ci).value = (
                        f'=IFERROR(({bcl}{r}/{bcl}{r-1}-1)*{pcl}{r},"")'
                    )
                else:
                    ws_cm.cell(r, 15 + ci).value = '=""'

            # GT total (col 21)
            if i >= 1:
                ws_cm.cell(r, 21).value = f'=IFERROR(SUM(O{r}:T{r})*100,"")'
            else:
                ws_cm.cell(r, 21).value = '=""'

            # Contrib GA mensuel (cols 23-28) — 12 mois
            for ci in range(6):
                bcl = get_column_letter(2 + ci)
                pcl = get_column_letter(8 + ci)
                if i >= 12:
                    ws_cm.cell(r, 23 + ci).value = (
                        f'=IFERROR(({bcl}{r}/{bcl}{r-12}-1)*{pcl}{r},"")'
                    )
                else:
                    ws_cm.cell(r, 23 + ci).value = '=""'

            # GA total (col 29)
            if i >= 12:
                ws_cm.cell(r, 29).value = f'=IFERROR(SUM(W{r}:AB{r})*100,"")'
            else:
                ws_cm.cell(r, 29).value = '=""'

            for c in range(1, 30):
                _copy_cell_style(ws_cm.cell(ref_cm, c), ws_cm.cell(r, c))

        for r in range(cm_start + n_months, orig_cm_max + 1):
            for c in range(1, 30):
                ws_cm.cell(r, c).value = None

    # ── 5) Contributions_Trim ─────────────────────────────────────────────
    if "Contributions_Trim" in wb.sheetnames:
        ws_ct = wb["Contributions_Trim"]
        n_q = len(df_quarterly)
        ct_start = 4
        orig_ct_max = ws_ct.max_row
        ref_ct = min(orig_ct_max, ct_start + 1)

        for i in range(n_q):
            r = ct_start + i
            trim_label = str(df_quarterly.iloc[i].get("Trimestre", ""))
            ws_ct.cell(r, 1).value = trim_label

            # ICAE rebasé (cols 2-7) = ICAE_Trimestriel cols J-O
            for ci in range(6):
                ws_ct.cell(r, 2 + ci).value = (
                    f"=ICAE_Trimestriel!{get_column_letter(10 + ci)}{r}"
                )

            # Poids PIB (cols 8-13)
            for ci in range(6):
                bcl = get_column_letter(2 + ci)
                poids_parts = "+".join(
                    f"IF(ISNUMBER({get_column_letter(2+cj)}{r}),Poids_PIB!$C${4+cj},0)"
                    for cj in range(6)
                )
                ws_ct.cell(r, 8 + ci).value = (
                    f"=IFERROR(IF(ISNUMBER({bcl}{r}),"
                    f"Poids_PIB!$C${4+ci}/({poids_parts}),0),0)"
                )

            # Contrib GT Trim (cols 15-20) — 1 trimestre lag
            for ci in range(6):
                bcl = get_column_letter(2 + ci)
                pcl = get_column_letter(8 + ci)
                if i >= 1:
                    ws_ct.cell(r, 15 + ci).value = (
                        f'=IFERROR(({bcl}{r}/{bcl}{r - 1}-1)*{pcl}{r},"")'
                    )
                else:
                    ws_ct.cell(r, 15 + ci).value = '=""'

            # GT total (col 21)
            if i >= 1:
                ws_ct.cell(r, 21).value = f'=IFERROR(SUM(O{r}:T{r})*100,"")'
            else:
                ws_ct.cell(r, 21).value = '=""'

            # Contrib GA Trim (cols 23-28) — 4 trimestres lag
            for ci in range(6):
                bcl = get_column_letter(2 + ci)
                pcl = get_column_letter(8 + ci)
                if i >= 4:
                    ws_ct.cell(r, 23 + ci).value = (
                        f'=IFERROR(({bcl}{r}/{bcl}{r - 4}-1)*{pcl}{r},"")'
                    )
                else:
                    ws_ct.cell(r, 23 + ci).value = '=""'

            # GA total (col 29)
            if i >= 4:
                ws_ct.cell(r, 29).value = f'=IFERROR(SUM(W{r}:AB{r})*100,"")'
            else:
                ws_ct.cell(r, 29).value = '=""'

            for c in range(1, 30):
                _copy_cell_style(ws_ct.cell(ref_ct, c), ws_ct.cell(r, c))

        for r in range(ct_start + n_q, orig_ct_max + 1):
            for c in range(1, 30):
                ws_ct.cell(r, c).value = None

    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    buf.seek(0)
    return buf.getvalue()


def _write_cemac_from_scratch(df_monthly: pd.DataFrame,
                               df_quarterly: pd.DataFrame,
                               poids: dict) -> bytes:
    """Fallback : génère un classeur CEMAC de zéro."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        poids_df = pd.DataFrame([
            {"Code": k, "Poids": v} for k, v in poids.items()
        ])
        poids_df.to_excel(writer, sheet_name="Poids_PIB", index=False)
        df_monthly.to_excel(writer, sheet_name="ICAE_Pays", index=True)
        df_quarterly.to_excel(writer, sheet_name="ICAE_Trimestriel",
                              index=False)
    buf.seek(0)
    return buf.getvalue()
