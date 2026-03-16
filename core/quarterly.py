"""Trimestrialisation et contributions sectorielles."""
import numpy as np
import pandas as pd


def quarterly_mean(monthly_series: pd.Series,
                   dates: pd.Series) -> pd.DataFrame:
    """Calcule l'ICAE trimestriel = moyenne des 3 mois du trimestre."""
    df = pd.DataFrame({"date": pd.to_datetime(dates), "value": monthly_series.values})
    df["quarter"] = df["date"].dt.to_period("Q")
    q = df.groupby("quarter").agg(
        icae_trim=("value", "mean"),
        debut=("date", "min"),
        fin=("date", "max"),
        n=("value", "count"),
    ).reset_index()
    q["trimestre"] = q["quarter"].astype(str)
    return q


def calc_ga_trim(icae_trim: pd.Series) -> pd.Series:
    """GA trimestriel = (ICAE_T - ICAE_{T-4}) / ICAE_{T-4} * 100."""
    return (icae_trim / icae_trim.shift(4) - 1) * 100


def calc_gt_trim(icae_trim: pd.Series) -> pd.Series:
    """GT trimestriel = (ICAE_T - ICAE_{T-1}) / ICAE_{T-1} * 100."""
    return (icae_trim / icae_trim.shift(1) - 1) * 100


def contributions_sectorielles(m_df: pd.DataFrame, codification: pd.DataFrame,
                               dates: pd.Series) -> pd.DataFrame:
    """
    Calcule les contributions sectorielles mensuelles puis trimestrielles.
    m_df : contributions par variable (m_{j,t})
    codification : doit contenir 'Code' et 'Secteur'
    """
    code_to_sector = dict(zip(codification["Code"], codification["Secteur"]))
    sectors = {}
    for col in m_df.columns:
        sect = code_to_sector.get(col, "Autre")
        if sect not in sectors:
            sectors[sect] = []
        sectors[sect].append(col)

    contrib = pd.DataFrame(index=m_df.index)
    contrib["Date"] = pd.to_datetime(dates).values
    for sect, cols in sectors.items():
        contrib[sect] = m_df[cols].sum(axis=1)

    return contrib


def contributions_sectorielles_trim(m_df: pd.DataFrame,
                                     codification: pd.DataFrame,
                                     dates: pd.Series) -> pd.DataFrame:
    """
    Calcule les contributions sectorielles trimestrielles.
    Retourne un DataFrame avec colonnes = secteurs et 'trimestre', 'GA_Trim'.
    """
    contrib_m = contributions_sectorielles(m_df, codification, dates)
    date_col = contrib_m["Date"]
    sector_cols = [c for c in contrib_m.columns if c != "Date"]

    contrib_m["_quarter"] = pd.to_datetime(date_col).dt.to_period("Q")
    q = contrib_m.groupby("_quarter")[sector_cols].mean().reset_index()
    q["trimestre"] = q["_quarter"].astype(str)
    q = q.drop(columns=["_quarter"])
    return q


def agg_m_to_q(monthly_df: pd.DataFrame, dates: pd.Series,
               agg_types: dict = None) -> pd.DataFrame:
    """
    Agrège des séries mensuelles en trimestriel.
    agg_types : dict {col: 'stock'|'flow'|'mean'}, défaut='mean'
    - stock : dernière valeur non-NA du trimestre
    - flow : somme
    - mean : moyenne
    """
    if agg_types is None:
        agg_types = {}

    df = monthly_df.copy()
    df["_date"] = pd.to_datetime(dates).values
    df["_quarter"] = df["_date"].dt.to_period("Q")

    cols = [c for c in df.columns if not c.startswith("_")]
    result_frames = []

    for col in cols:
        agg_type = agg_types.get(col, "mean")
        if agg_type == "stock":
            q_series = df.groupby("_quarter")[col].apply(
                lambda s: s.dropna().iloc[-1] if s.dropna().shape[0] > 0 else np.nan
            )
        elif agg_type == "flow":
            q_series = df.groupby("_quarter")[col].sum(min_count=1)
        else:
            q_series = df.groupby("_quarter")[col].mean()
        result_frames.append(q_series.rename(col))

    result = pd.concat(result_frames, axis=1)
    result.index.name = "quarter"
    return result.reset_index()


def normalize_contrib_to_ga(contrib_trim: pd.DataFrame,
                            ga_trim: pd.Series) -> pd.DataFrame:
    """
    Normalise les contributions sectorielles trimestrielles pour que
    leur somme soit égale au GA trimestriel à chaque période.
    Cela permet d'afficher barres empilées + courbe GA sur le même axe (%).
    """
    sector_cols = [c for c in contrib_trim.columns
                   if c not in ("trimestre", "Date", "_quarter")]
    result = contrib_trim.copy()
    total = result[sector_cols].sum(axis=1)
    for i in range(len(result)):
        t = total.iloc[i]
        ga = ga_trim.iloc[i] if i < len(ga_trim) else np.nan
        if pd.notna(ga) and pd.notna(t) and abs(t) > 1e-10:
            scale = ga / t
            for col in sector_cols:
                result.iloc[i, result.columns.get_loc(col)] *= scale
    return result
