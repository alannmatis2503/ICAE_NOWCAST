"""Moteur d'agrégation CEMAC."""
import numpy as np
import pandas as pd
from config import POIDS_PIB, COUNTRY_CODES, COUNTRY_NAMES


def compute_icae_cemac(icae_dict: dict, poids: dict = None) -> pd.DataFrame:
    """
    Calcule l'ICAE CEMAC agrégé.

    Parameters
    ----------
    icae_dict : {code_pays: pd.Series} — ICAE mensuel par pays
    poids : {code_pays: float} — pondérations PIB (default = PIB 2014)

    Returns
    -------
    DataFrame avec colonnes: Date, + pays, ICAE_CEMAC, GA, GT
    """
    if poids is None:
        poids = POIDS_PIB

    # Aligner toutes les séries sur un index commun
    all_series = {}
    for code in COUNTRY_CODES:
        if code in icae_dict and icae_dict[code] is not None:
            s = icae_dict[code].copy()
            s.name = code
            all_series[code] = s

    if not all_series:
        return pd.DataFrame()

    df = pd.DataFrame(all_series)

    # ICAE CEMAC = somme pondérée (seuls les pays avec données contribuent)
    country_cols = list(all_series.keys())
    weights_series = pd.Series({code: poids.get(code, 0) for code in country_cols})

    # Pour chaque ligne, calculer la moyenne pondérée avec les pays disponibles
    icae_cemac = pd.Series(np.nan, index=df.index)
    for idx in df.index:
        row = df.loc[idx, country_cols]
        available = row.dropna()
        if len(available) == 0:
            continue
        w = weights_series[available.index]
        w_sum = w.sum()
        if w_sum > 0:
            icae_cemac[idx] = (available * w).sum() / w_sum

    df["ICAE_CEMAC"] = icae_cemac
    df["GA"] = (df["ICAE_CEMAC"] / df["ICAE_CEMAC"].shift(12) - 1) * 100
    df["GT"] = (df["ICAE_CEMAC"] / df["ICAE_CEMAC"].shift(1) - 1) * 100

    return df


def quarterly_cemac(df_monthly: pd.DataFrame,
                    dates: pd.Series) -> pd.DataFrame:
    """Trimestrialise l'ICAE CEMAC."""
    df = df_monthly.copy()
    df["_date"] = pd.to_datetime(dates).values
    df["_quarter"] = df["_date"].dt.to_period("Q")

    cols = [c for c in df.columns if not c.startswith("_")]
    result = df.groupby("_quarter")[cols].mean()
    result = result.reset_index()
    result.rename(columns={"_quarter": "Trimestre"}, inplace=True)

    # GA trimestriel
    if "ICAE_CEMAC" in result.columns:
        result["GA_Trim"] = (
            result["ICAE_CEMAC"] / result["ICAE_CEMAC"].shift(4) - 1
        ) * 100

    return result
