"""Moteur de calcul ICAE — réimplémentation fidèle du pipeline R Shiny."""
import numpy as np
import pandas as pd
from config import SYM_FACTOR, I_INIT, DENOM_FLOOR, SIGMA_FLOOR, SUM_M_CAP


# ────────────────────────────────────────────────────────────────────────────
# 1. Taux de Croissance Symétrique (TCS)
# ────────────────────────────────────────────────────────────────────────────
def calc_sym_growth(series: pd.Series) -> pd.Series:
    """C_t = 200 * (X_t - X_{t-1}) / (X_t + X_{t-1})"""
    x = series.astype(float)
    x_prev = x.shift(1)
    denom = x + x_prev
    denom = denom.where(denom.abs() >= DENOM_FLOOR, np.nan)
    return SYM_FACTOR * (x - x_prev) / denom


def calc_sym_growth_df(df: pd.DataFrame) -> pd.DataFrame:
    """Applique le TCS à toutes les colonnes d'un DataFrame."""
    return df.apply(calc_sym_growth)


# ────────────────────────────────────────────────────────────────────────────
# 2. Écart-type (fixe ou glissant)
# ────────────────────────────────────────────────────────────────────────────
def fixed_sigma(tcs_df: pd.DataFrame, base_rows: range) -> pd.Series:
    """σ_j = STDEV(TCS_j sur l'année de base). Retourne un σ par variable."""
    subset = tcs_df.iloc[base_rows]
    sigma = subset.std(ddof=1)
    return sigma.clip(lower=SIGMA_FLOOR)


def rolling_sigma(tcs_df: pd.DataFrame, window: int = 12,
                  min_obs: int = 4) -> pd.DataFrame:
    """σ_{j,t} = rolling std sur les `window` mois précédents (excluant t)."""
    shifted = tcs_df.shift(1)
    sigma = shifted.rolling(window=window, min_periods=min_obs).std(ddof=1)
    return sigma.clip(lower=SIGMA_FLOOR)


# ────────────────────────────────────────────────────────────────────────────
# 3. Pondérations & signal filtré
# ────────────────────────────────────────────────────────────────────────────
def calc_weights_fixed(tcs_df: pd.DataFrame, sigma: pd.Series,
                       priors: pd.Series) -> dict:
    """Mode écart-type fixe : pondérations constantes dans le temps."""
    omega = 1.0 / sigma                             # ω_j = 1/σ_j
    a = omega / omega.sum()                          # a_j normalisé
    ab = a * priors                                  # a×b
    pond_finale = ab / ab.sum()                      # pondération finale
    m = tcs_df.multiply(pond_finale, axis=1)         # contributions
    sum_m = m.sum(axis=1)                            # signal composite
    return {
        "omega": omega, "a": a, "pond_finale": pond_finale,
        "m": m, "sum_m": sum_m,
    }


def calc_weights_rolling(tcs_df: pd.DataFrame, sigma_df: pd.DataFrame,
                         priors: pd.Series) -> dict:
    """Mode écart-type glissant : pondérations varient dans le temps."""
    omega_df = 1.0 / sigma_df
    # Pondérer par les priors
    omega_rho = omega_df.multiply(priors, axis=1)
    # Normaliser ligne par ligne
    row_sums = omega_rho.sum(axis=1)
    sf = omega_rho.div(row_sums, axis=0)
    m = tcs_df * sf
    sum_m = m.sum(axis=1)
    return {
        "omega": omega_df, "sf": sf, "pond_finale": sf,
        "m": m, "sum_m": sum_m,
    }


# ────────────────────────────────────────────────────────────────────────────
# 4. Indice récursif
# ────────────────────────────────────────────────────────────────────────────
def calc_I_recursive(sum_m: pd.Series, I0: float = I_INIT,
                     start_idx: int = 0) -> pd.Series:
    """I_t = I_{t-1} * (200 + Σm_t) / (200 - Σm_t)"""
    I = pd.Series(np.nan, index=sum_m.index, dtype=float)
    I.iloc[start_idx] = I0
    for t in range(start_idx + 1, len(sum_m)):
        sm = sum_m.iloc[t]
        if pd.isna(sm) or pd.isna(I.iloc[t - 1]):
            continue
        if abs(sm) >= SUM_M_CAP:
            continue
        denom = SYM_FACTOR - sm
        if abs(denom) < DENOM_FLOOR:
            continue
        I.iloc[t] = I.iloc[t - 1] * (SYM_FACTOR + sm) / denom
    return I


# ────────────────────────────────────────────────────────────────────────────
# 5. Normalisation base 100
# ────────────────────────────────────────────────────────────────────────────
def normalize_base100(I: pd.Series, dates: pd.Series,
                      base_year: int) -> pd.Series:
    """ICAE_t = 100 * I_t / mean(I sur l'année de base)"""
    mask = dates.dt.year == base_year
    I_base = I[mask].dropna()
    if len(I_base) < 3:
        raise ValueError(
            f"Pas assez d'observations ({len(I_base)}) pour l'année de base {base_year}"
        )
    B0 = I_base.mean()
    return 100.0 * I / B0


# ────────────────────────────────────────────────────────────────────────────
# 6. Pipeline complet
# ────────────────────────────────────────────────────────────────────────────
def run_icae_pipeline(donnees: pd.DataFrame, priors: pd.Series,
                      base_year: int, base_rows: range,
                      sigma_mode: str = "fixed",
                      rolling_window: int = 12) -> dict:
    """
    Pipeline complet ICAE.

    Parameters
    ----------
    donnees : DataFrame avec colonne 'Date' + colonnes variables
    priors : Series indexée par code variable (scores PRIOR médians)
    base_year : année de base (ex: 2023)
    base_rows : range des lignes dans le TCS correspondant à l'année de base
    sigma_mode : "fixed" ou "rolling"
    rolling_window : fenêtre si mode rolling

    Returns
    -------
    dict avec toutes les étapes intermédiaires et le résultat final
    """
    dates = pd.to_datetime(donnees["Date"])
    data_cols = [c for c in donnees.columns if c != "Date"]
    df = donnees[data_cols].astype(float)

    # Aligner les priors sur les colonnes
    priors_aligned = priors.reindex(data_cols).fillna(0)
    active_mask = priors_aligned > 0
    active_cols = priors_aligned[active_mask].index.tolist()

    # Étape 2 : TCS
    tcs = calc_sym_growth_df(df)

    # Colonnes actives seulement pour le calcul
    tcs_active = tcs[active_cols]
    priors_active = priors_aligned[active_cols]

    # Étape 3 : Pondérations
    if sigma_mode == "fixed":
        sigma = fixed_sigma(tcs_active, base_rows)
        weights = calc_weights_fixed(tcs_active, sigma, priors_active)
    else:
        sigma_df = rolling_sigma(tcs_active, window=rolling_window)
        weights = calc_weights_rolling(tcs_active, sigma_df, priors_active)

    sum_m = weights["sum_m"]

    # Trouver la première valeur non-NaN de sum_m pour démarrer
    first_valid = sum_m.first_valid_index()
    if first_valid is None:
        raise ValueError("Aucune valeur valide dans le signal composite")
    start_pos = sum_m.index.get_loc(first_valid)

    # Étape 4 : Indice récursif
    I = calc_I_recursive(sum_m, I0=I_INIT, start_idx=start_pos)

    # Étape 5 : Normalisation
    icae = normalize_base100(I, dates, base_year)

    # GA mensuel
    ga_monthly = (icae / icae.shift(12) - 1) * 100

    return {
        "dates": dates,
        "tcs": tcs,
        "tcs_active": tcs_active,
        "weights": weights,
        "sum_m": sum_m,
        "indice": I,
        "icae": icae,
        "ga_monthly": ga_monthly,
        "active_cols": active_cols,
        "priors": priors_active,
        "pond_finale": weights["pond_finale"],
        "base_year": base_year,
    }
