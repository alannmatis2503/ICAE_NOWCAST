"""Moteur Nowcast — Bridge, U-MIDAS, PC, DFM-lite."""
import numpy as np
import pandas as pd
import warnings
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler


def _safe_ols(X: np.ndarray, y: np.ndarray):
    """OLS avec protection contre les matrices singulières."""
    try:
        from numpy.linalg import lstsq
        beta, _, _, _ = lstsq(X, y, rcond=None)
        return beta
    except Exception:
        return None


def _prepare_data(pib_q: pd.Series, hf_q: pd.DataFrame,
                  h_ahead: int = 4) -> dict:
    """Prépare les données communes pour tous les modèles."""
    # Aligner PIB et HF par quarter
    common = pib_q.dropna().index.intersection(hf_q.index)
    if len(common) < 10:
        return None

    pib = pib_q.loc[common].values.astype(float)
    hf = hf_q.loc[common].copy()

    # Supprimer les colonnes constantes
    hf = hf.loc[:, hf.std() > 0]
    if hf.shape[1] == 0:
        return None

    # Imputer les NA par la moyenne
    hf = hf.fillna(hf.mean())

    n = len(pib)
    n_train = max(n - h_ahead, 10)

    return {"pib": pib, "hf": hf, "n": n, "n_train": n_train,
            "index": common}


# ────────────────────────────────────────────────────────────────────────────
# Bridge Model
# ────────────────────────────────────────────────────────────────────────────
def fit_bridge(pib_q: pd.Series, hf_q: pd.DataFrame,
               h_ahead: int = 4) -> dict:
    """Bridge : PC1 + 2 meilleurs indicateurs (retardés)."""
    data = _prepare_data(pib_q, hf_q, h_ahead)
    if data is None:
        return {"forecast": pd.Series(dtype=float), "name": "Bridge"}

    pib, hf, n, n_train = data["pib"], data["hf"], data["n"], data["n_train"]

    # PCA
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(hf.values)
    pca = PCA(n_components=min(3, X_scaled.shape[1]))
    pcs = pca.fit_transform(X_scaled)
    pc1 = pcs[:, 0]

    # Top 2 corrélés
    corrs = {}
    for col in hf.columns:
        c = np.corrcoef(pib[:n_train], hf[col].values[:n_train])[0, 1]
        if not np.isnan(c):
            corrs[col] = abs(c)
    top2 = sorted(corrs, key=corrs.get, reverse=True)[:2]

    # Construire les régresseurs avec retards
    n_eff = n - 1  # on perd 1 obs par le lag
    X = np.column_stack([
        np.ones(n_eff),
        pc1[1:],                    # PC1 contemporain
        pc1[:-1],                   # PC1 lag 1
    ])
    for v in top2:
        X = np.column_stack([X, hf[v].values[:-1]])  # lag 1

    y = pib[1:]
    n_train_eff = n_train - 1

    beta = _safe_ols(X[:n_train_eff], y[:n_train_eff])
    if beta is None:
        return {"forecast": pd.Series(dtype=float), "name": "Bridge"}

    forecast = X @ beta
    result = pd.Series(np.nan, index=data["index"])
    result.iloc[1:] = forecast

    return {"forecast": result, "name": "Bridge", "beta": beta}


# ────────────────────────────────────────────────────────────────────────────
# U-MIDAS
# ────────────────────────────────────────────────────────────────────────────
def fit_umidas(pib_q: pd.Series, hf_q: pd.DataFrame,
               h_ahead: int = 4) -> dict:
    """U-MIDAS : meilleur HF + 2 retards."""
    data = _prepare_data(pib_q, hf_q, h_ahead)
    if data is None:
        return {"forecast": pd.Series(dtype=float), "name": "U-MIDAS"}

    pib, hf, n, n_train = data["pib"], data["hf"], data["n"], data["n_train"]

    # Meilleur indicateur
    corrs = {}
    for col in hf.columns:
        c = np.corrcoef(pib[:n_train], hf[col].values[:n_train])[0, 1]
        if not np.isnan(c):
            corrs[col] = abs(c)
    if not corrs:
        return {"forecast": pd.Series(dtype=float), "name": "U-MIDAS"}
    best = max(corrs, key=corrs.get)
    x = hf[best].values

    n_eff = n - 2
    X = np.column_stack([
        np.ones(n_eff),
        x[2:],      # contemporain
        x[1:-1],    # lag 1
        x[:-2],     # lag 2
    ])
    y = pib[2:]
    n_train_eff = n_train - 2

    beta = _safe_ols(X[:n_train_eff], y[:n_train_eff])
    if beta is None:
        return {"forecast": pd.Series(dtype=float), "name": "U-MIDAS"}

    forecast = X @ beta
    result = pd.Series(np.nan, index=data["index"])
    result.iloc[2:] = forecast

    return {"forecast": result, "name": "U-MIDAS"}


# ────────────────────────────────────────────────────────────────────────────
# PC Regression
# ────────────────────────────────────────────────────────────────────────────
def fit_pc(pib_q: pd.Series, hf_q: pd.DataFrame,
           h_ahead: int = 4, n_components: int = 3) -> dict:
    """Régression sur les composantes principales."""
    data = _prepare_data(pib_q, hf_q, h_ahead)
    if data is None:
        return {"forecast": pd.Series(dtype=float), "name": "PC"}

    pib, hf, n, n_train = data["pib"], data["hf"], data["n"], data["n_train"]

    k = min(n_components, hf.shape[1], n_train - 2)
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(hf.values)
    pca = PCA(n_components=k)
    pcs = pca.fit_transform(X_scaled)

    X = np.column_stack([np.ones(n), pcs])
    beta = _safe_ols(X[:n_train], pib[:n_train])
    if beta is None:
        return {"forecast": pd.Series(dtype=float), "name": "PC"}

    forecast = X @ beta
    return {"forecast": pd.Series(forecast, index=data["index"]),
            "name": "PC"}


# ────────────────────────────────────────────────────────────────────────────
# DFM-lite
# ────────────────────────────────────────────────────────────────────────────
def fit_dfm(pib_q: pd.Series, hf_q: pd.DataFrame,
            h_ahead: int = 4) -> dict:
    """DFM-lite : PC1 + lag(PC1) → PIB."""
    data = _prepare_data(pib_q, hf_q, h_ahead)
    if data is None:
        return {"forecast": pd.Series(dtype=float), "name": "DFM"}

    pib, hf, n, n_train = data["pib"], data["hf"], data["n"], data["n_train"]

    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(hf.values)
    pca = PCA(n_components=1)
    f = pca.fit_transform(X_scaled).flatten()

    n_eff = n - 1
    X = np.column_stack([
        np.ones(n_eff),
        f[1:],     # contemporain
        f[:-1],    # lag 1
    ])
    y = pib[1:]
    n_train_eff = n_train - 1

    beta = _safe_ols(X[:n_train_eff], y[:n_train_eff])
    if beta is None:
        return {"forecast": pd.Series(dtype=float), "name": "DFM"}

    forecast = X @ beta
    result = pd.Series(np.nan, index=data["index"])
    result.iloc[1:] = forecast

    return {"forecast": result, "name": "DFM"}


# ────────────────────────────────────────────────────────────────────────────
# Métriques
# ────────────────────────────────────────────────────────────────────────────
def compute_metrics(actual: np.ndarray, predicted: np.ndarray) -> dict:
    """RMSE, MAE, MAPE."""
    mask = ~(np.isnan(actual) | np.isnan(predicted))
    a, p = actual[mask], predicted[mask]
    if len(a) == 0:
        return {"rmse": np.nan, "mae": np.nan, "mape": np.nan}
    rmse = np.sqrt(np.mean((a - p) ** 2))
    mae = np.mean(np.abs(a - p))
    mape = np.mean(np.abs((a - p) / a)) * 100 if np.all(a != 0) else np.nan
    return {"rmse": rmse, "mae": mae, "mape": mape}


def compute_ins_out_metrics(actual: pd.Series, predicted: pd.Series,
                            h_test: int = 3) -> dict:
    """Métriques in-sample et out-of-sample."""
    common = actual.dropna().index.intersection(predicted.dropna().index)
    if len(common) < h_test + 2:
        return {"in_sample": {}, "out_sample": {}}
    a = actual.loc[common].values
    p = predicted.loc[common].values
    split = len(a) - h_test
    return {
        "in_sample": compute_metrics(a[:split], p[:split]),
        "out_sample": compute_metrics(a[split:], p[split:]),
    }


# ────────────────────────────────────────────────────────────────────────────
# Pipeline complet
# ────────────────────────────────────────────────────────────────────────────
MODEL_DISPATCH = {
    "Bridge": fit_bridge,
    "U-MIDAS": fit_umidas,
    "PC": fit_pc,
    "DFM": fit_dfm,
}


def run_nowcast(pib_q: pd.Series, hf_q: pd.DataFrame,
                models: list = None, h_ahead: int = 4,
                n_components: int = 3, h_test: int = 3) -> dict:
    """
    Lance tous les modèles Nowcast.

    Returns : dict {model_name: {forecast, metrics_in, metrics_out}}
    """
    if models is None:
        models = list(MODEL_DISPATCH.keys())

    results = {}
    for name in models:
        fn = MODEL_DISPATCH[name]
        if name == "PC":
            res = fn(pib_q, hf_q, h_ahead, n_components)
        else:
            res = fn(pib_q, hf_q, h_ahead)

        metrics = compute_ins_out_metrics(pib_q, res["forecast"], h_test)
        results[name] = {
            "forecast": res["forecast"],
            "metrics": metrics,
        }

    # Corrélations
    for name, r in results.items():
        fc = r["forecast"]
        common = pib_q.dropna().index.intersection(fc.dropna().index)
        if len(common) >= 3:
            corr = np.corrcoef(pib_q.loc[common], fc.loc[common])[0, 1]
        else:
            corr = np.nan
        r["correlation"] = corr

    return results
