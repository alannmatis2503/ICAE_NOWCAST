"""
Désagrégation temporelle : annuel → trimestriel ou mensuel.

Méthodes :
  1. Chow-Lin (maximum log-likelihood) avec indicateur HF comme régresseur
  2. Fallback : Denton-Cholette (sans régresseur)
  3. Recalage : sum(sous-périodes) = valeur annuelle
"""
import numpy as np
import pandas as pd
from scipy.optimize import minimize_scalar
from scipy.linalg import cho_factor, cho_solve

# Nombre de sous-périodes par fréquence cible
_FREQ_S = {"Trimestrielle": 4, "Mensuelle": 12}


# ────────────────────────────────────────────────────────────────────────────
# Matrice d'agrégation C (somme s sous-périodes → annuel)
# ────────────────────────────────────────────────────────────────────────────
def _build_C(n_sub: int, n_years: int, s: int = 4) -> np.ndarray:
    """Matrice C (n_years × n_sub) telle que C @ y_hf = y_a."""
    C = np.zeros((n_years, n_sub))
    for i in range(n_years):
        C[i, i * s: i * s + s] = 1.0
    return C


# ────────────────────────────────────────────────────────────────────────────
# Chow-Lin max-log-likelihood
# ────────────────────────────────────────────────────────────────────────────
def _ar1_vcov(n: int, rho: float) -> np.ndarray:
    """Matrice de variance-covariance AR(1) : Σ_{ij} = rho^|i-j|."""
    idx = np.arange(n)
    return rho ** np.abs(idx[:, None] - idx[None, :])


def _chow_lin_objective(rho: float, y_a: np.ndarray, X_hf: np.ndarray,
                        C: np.ndarray) -> float:
    """Negative log-likelihood concentrée pour Chow-Lin."""
    n_q = X_hf.shape[0]
    if abs(rho) >= 0.999:
        return 1e15
    Sigma_q = _ar1_vcov(n_q, rho)
    Sigma_a = C @ Sigma_q @ C.T
    try:
        L, low = cho_factor(Sigma_a)
        X_a = C @ X_hf
        beta = np.linalg.lstsq(X_a, y_a, rcond=None)[0]
        u_a = y_a - X_a @ beta
        val = cho_solve((L, low), u_a)
        n_a = len(y_a)
        log_det = 2 * np.sum(np.log(np.diag(L)))
        nll = 0.5 * (n_a * np.log(2 * np.pi) + log_det + u_a @ val)
        return nll
    except Exception:
        return 1e15


def chow_lin(y_a: np.ndarray, X_hf: np.ndarray, s: int = 4) -> np.ndarray:
    """
    Chow-Lin maxlog : désagrège y_a (annuel) en sous-périodes
    en utilisant X_hf (régresseur haute fréquence).

    Parameters
    ----------
    s : nombre de sous-périodes par an (4=trim, 12=mensuel)

    Returns : y_hf estimé (n_sub,)
    """
    n_a = len(y_a)
    n_hf = X_hf.shape[0]
    C = _build_C(n_hf, n_a, s)

    res = minimize_scalar(
        _chow_lin_objective, bounds=(0.01, 0.99), method="bounded",
        args=(y_a, X_hf, C),
    )
    rho_opt = res.x

    Sigma_q = _ar1_vcov(n_hf, rho_opt)
    Sigma_a = C @ Sigma_q @ C.T
    Sigma_a_inv = np.linalg.inv(Sigma_a)

    X_a = C @ X_hf
    beta = np.linalg.lstsq(X_a, y_a, rcond=None)[0]

    y_hf_prelim = X_hf @ beta
    u_a = y_a - C @ y_hf_prelim

    y_hf = y_hf_prelim + Sigma_q @ C.T @ Sigma_a_inv @ u_a
    return y_hf


# ────────────────────────────────────────────────────────────────────────────
# Denton-Cholette (sans régresseur — fallback)
# ────────────────────────────────────────────────────────────────────────────
def denton_cholette(y_a: np.ndarray, n_sub: int, s: int = 4) -> np.ndarray:
    """
    Denton-Cholette proportionnel (sans indicateur).
    Distribue la valeur annuelle uniformément sur les s sous-périodes.
    """
    n_a = len(y_a)
    y_hf = np.zeros(n_sub)
    for i in range(n_a):
        y_hf[i * s: i * s + s] = y_a[i] / s
    return y_hf


# ────────────────────────────────────────────────────────────────────────────
# Ecotrim / Fernandez (sans régresseur ou avec)
# ────────────────────────────────────────────────────────────────────────────
def _diff_matrix(n: int) -> np.ndarray:
    """Matrice de différences premières D (n-1 × n)."""
    D = np.zeros((n - 1, n))
    for i in range(n - 1):
        D[i, i] = -1.0
        D[i, i + 1] = 1.0
    return D


def fernandez(y_a: np.ndarray, X_hf: np.ndarray | None = None,
              s: int = 4) -> np.ndarray:
    """
    Méthode de Fernandez (1981) — variante GLS sans estimation de rho.
    Utilisée dans Ecotrim. Minimise les différences premières des résidus
    sous la contrainte d'agrégation temporelle.
    """
    n_a = len(y_a)
    n_hf = n_a * s
    C = _build_C(n_hf, n_a, s)
    D = _diff_matrix(n_hf)
    # Sigma = (D'D)^{-1}  (matrice de lissage)
    DtD = D.T @ D
    DtD_inv = np.linalg.inv(DtD + 1e-8 * np.eye(n_hf))

    if X_hf is not None and X_hf.shape[0] >= n_hf:
        X = X_hf[:n_hf]
        X_a = C @ X
        beta = np.linalg.lstsq(X_a, y_a, rcond=None)[0]
        p_hf = X @ beta
    else:
        p_hf = np.zeros(n_hf)

    u_a = y_a - C @ p_hf
    Sigma_a = C @ DtD_inv @ C.T
    Sigma_a_inv = np.linalg.inv(Sigma_a + 1e-8 * np.eye(n_a))
    y_hf = p_hf + DtD_inv @ C.T @ Sigma_a_inv @ u_a
    return y_hf


# ────────────────────────────────────────────────────────────────────────────
# Pipeline complet
# ────────────────────────────────────────────────────────────────────────────
def disaggregate_annual(
    y_annual: pd.Series,
    target_freq: str = "Trimestrielle",
    hf_indicator: pd.DataFrame = None,
    method: str | None = None,
) -> pd.Series:
    """
    Désagrège une série annuelle vers la fréquence cible.

    Parameters
    ----------
    y_annual : Series indexée par année (int ou convertible)
    target_freq : "Trimestrielle" ou "Mensuelle"
    hf_indicator : DataFrame d'indicateurs HF (optionnel)
    method : "chow-lin", "denton", "ecotrim", ou None (auto)

    Returns
    -------
    Series indexée par PeriodIndex (Q ou M)
    """
    s = _FREQ_S.get(target_freq, 4)
    freq_code = "Q" if s == 4 else "M"

    years = sorted(y_annual.dropna().index.astype(int))
    n_a = len(years)
    n_hf = n_a * s
    y_a = y_annual.loc[years].values.astype(float)

    # Préparer indicateur HF si disponible
    X_hf_matrix = None
    if hf_indicator is not None and hf_indicator.shape[1] > 0:
        from sklearn.decomposition import PCA
        from sklearn.preprocessing import StandardScaler

        hf = hf_indicator.dropna(how="all", axis=1)
        hf = hf.loc[:, hf.std() > 0]

        if hf.shape[1] > 0 and hf.dropna().shape[0] >= n_hf:
            hf_filled = hf.iloc[:n_hf].fillna(hf.mean())
            scaler = StandardScaler()
            X_scaled = scaler.fit_transform(hf_filled)
            pca = PCA(n_components=1)
            pc1 = pca.fit_transform(X_scaled).flatten()
            X_hf_matrix = np.column_stack([np.ones(n_hf), pc1])

    # Déterminer la méthode
    if method is None:
        method = "chow-lin" if X_hf_matrix is not None else "denton"

    y_hf = None
    if method == "chow-lin" and X_hf_matrix is not None:
        try:
            y_hf = chow_lin(y_a, X_hf_matrix, s=s)
        except Exception:
            pass
    elif method == "ecotrim":
        try:
            y_hf = fernandez(y_a, X_hf_matrix, s=s)
        except Exception:
            pass

    if y_hf is None:
        y_hf = denton_cholette(y_a, n_hf, s=s)

    y_hf = _recalibrate(y_hf, y_a, n_a, s=s)
    idx = pd.period_range(
        start=f"{years[0]}Q1" if s == 4 else f"{years[0]}-01",
        periods=n_hf, freq=freq_code,
    )
    return pd.Series(y_hf, index=idx, name="disagg")


# Alias rétrocompatible
def disaggregate_annual_to_quarterly(
    pib_annual: pd.Series,
    hf_quarterly: pd.DataFrame = None,
    method: str | None = None,
) -> pd.Series:
    return disaggregate_annual(pib_annual, "Trimestrielle", hf_quarterly,
                               method=method)


def _recalibrate(y_hf: np.ndarray, y_a: np.ndarray, n_a: int,
                 s: int = 4) -> np.ndarray:
    """Recalage : sum(sous-périodes) = valeur annuelle pour chaque année."""
    y_hf = y_hf.copy()
    for i in range(n_a):
        total = y_hf[i * s: i * s + s].sum()
        if abs(total) > 1e-10:
            y_hf[i * s: i * s + s] *= y_a[i] / total
    return y_hf
