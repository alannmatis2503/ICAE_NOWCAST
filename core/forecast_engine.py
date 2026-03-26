"""Moteur de prévision — méthodes naïves + ARIMA/ETS."""
import numpy as np
import pandas as pd
import warnings

# Périodicité saisonnière par fréquence
SEASON_MAP = {"Mensuelle": 12, "Trimestrielle": 4, "Annuelle": 1}


def get_methods_for_frequency(freq: str) -> list:
    """Retourne les méthodes de prévision pertinentes pour la fréquence."""
    base = ["MM3", "MM6", "MM12", "TL"]
    if freq in ("Mensuelle", "Trimestrielle"):
        base += ["NS", "CS", "ARIMA", "ETS"]
    elif freq == "Annuelle":
        base = ["MM3", "TL", "ARIMA"]
    return base


# ────────────────────────────────────────────────────────────────────────────
# Méthodes naïves
# ────────────────────────────────────────────────────────────────────────────
def forecast_mm(series: pd.Series, h: int, window: int) -> np.ndarray:
    """Moyenne mobile sur les `window` dernières observations."""
    vals = series.dropna().values
    if len(vals) < window:
        return np.full(h, np.nan)
    base = vals[-window:].mean()
    return np.full(h, base)


def forecast_naive_seasonal(series: pd.Series, h: int,
                            season: int = 12) -> np.ndarray:
    """Naïve saisonnière : X_{t+h} = X_{t+h-season}."""
    vals = series.dropna().values
    preds = np.full(h, np.nan)
    for i in range(h):
        idx = len(vals) - season + i
        if 0 <= idx < len(vals):
            preds[i] = vals[idx]
        elif i >= season:
            preds[i] = preds[i - season]
    return preds


def forecast_seasonal_growth(series: pd.Series, h: int,
                             season: int = 12) -> np.ndarray:
    """Croissance saisonnière : X_{t+h} = X_{t+h-12} * X_t / X_{t-12}."""
    vals = series.dropna().values
    n = len(vals)
    if n < season + 1:
        return np.full(h, np.nan)
    growth = vals[-1] / vals[-1 - season] if vals[-1 - season] != 0 else 1.0
    preds = np.full(h, np.nan)
    for i in range(h):
        idx = n - season + i
        if 0 <= idx < n:
            preds[i] = vals[idx] * growth
    return preds


def forecast_trend_linear(series: pd.Series, h: int,
                          window: int = 24) -> np.ndarray:
    """Tendance linéaire sur les `window` derniers mois."""
    vals = series.dropna().values
    w = min(window, len(vals))
    y = vals[-w:]
    x = np.arange(w)
    if w < 2:
        return np.full(h, np.nan)
    coeffs = np.polyfit(x, y, 1)
    future_x = np.arange(w, w + h)
    return np.polyval(coeffs, future_x)


# ────────────────────────────────────────────────────────────────────────────
# ARIMA / ETS
# ────────────────────────────────────────────────────────────────────────────
def forecast_arima(series: pd.Series, h: int) -> dict:
    """Auto-ARIMA via pmdarima."""
    try:
        import pmdarima as pm
        vals = series.dropna().values
        if len(vals) < 24:
            return {"forecast": np.full(h, np.nan), "conf_int": None}
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            model = pm.auto_arima(
                vals, seasonal=True, m=12,
                stepwise=True, suppress_warnings=True,
                error_action="ignore",
                max_order=5, max_p=3, max_q=3,
            )
            fc, conf = model.predict(n_periods=h, return_conf_int=True,
                                     alpha=0.2)
        return {"forecast": fc, "conf_int": conf}
    except Exception:
        return {"forecast": np.full(h, np.nan), "conf_int": None}


def forecast_ets(series: pd.Series, h: int) -> dict:
    """ETS via statsmodels."""
    try:
        from statsmodels.tsa.holtwinters import ExponentialSmoothing
        vals = series.dropna().values
        if len(vals) < 24:
            return {"forecast": np.full(h, np.nan), "conf_int": None}
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            model = ExponentialSmoothing(
                vals, seasonal_periods=12,
                trend="add", seasonal="add",
                initialization_method="estimated",
            ).fit(optimized=True)
            fc = model.forecast(h)
            # IC approximatif
            residuals = model.resid
            se = np.std(residuals) * np.sqrt(np.arange(1, h + 1))
            lo = fc - 1.28 * se  # 80%
            hi = fc + 1.28 * se
        return {"forecast": fc, "conf_int": np.column_stack([lo, hi])}
    except Exception:
        return {"forecast": np.full(h, np.nan), "conf_int": None}


# ────────────────────────────────────────────────────────────────────────────
# Backtesting
# ────────────────────────────────────────────────────────────────────────────
def backtest_method(series: pd.Series, method_fn, h: int,
                    bt_window: int = 12, **kwargs) -> dict:
    """
    Backtesting par validation glissante.
    Retourne MAPE et RMSE sur la fenêtre de backtesting.
    """
    vals = series.dropna().values
    n = len(vals)
    if n < bt_window + h + 12:
        return {"mape": np.nan, "rmse": np.nan}

    errors = []
    for t in range(n - bt_window, n):
        train = pd.Series(vals[:t])
        actual = vals[t: t + 1]
        if len(actual) == 0:
            continue
        pred = method_fn(train, h=1, **kwargs)
        if isinstance(pred, dict):
            pred = pred["forecast"]
        if len(pred) > 0 and not np.isnan(pred[0]) and actual[0] != 0:
            errors.append((actual[0], pred[0]))

    if not errors:
        return {"mape": np.nan, "rmse": np.nan}

    actuals = np.array([e[0] for e in errors])
    preds = np.array([e[1] for e in errors])
    mape = np.mean(np.abs((actuals - preds) / actuals)) * 100
    rmse = np.sqrt(np.mean((actuals - preds) ** 2))
    return {"mape": mape, "rmse": rmse}


# ────────────────────────────────────────────────────────────────────────────
# Dispatch
# ────────────────────────────────────────────────────────────────────────────
METHOD_DISPATCH = {
    "MM3":  lambda s, h: forecast_mm(s, h, 3),
    "MM6":  lambda s, h: forecast_mm(s, h, 6),
    "MM12": lambda s, h: forecast_mm(s, h, 12),
    "NS":   lambda s, h: forecast_naive_seasonal(s, h),
    "CS":   lambda s, h: forecast_seasonal_growth(s, h),
    "TL":   lambda s, h: forecast_trend_linear(s, h),
    "ARIMA": lambda s, h: forecast_arima(s, h),
    "ETS":  lambda s, h: forecast_ets(s, h),
}

BACKTEST_DISPATCH = {
    "MM3":  lambda s, h, bt: backtest_method(s, forecast_mm, h, bt, window=3),
    "MM6":  lambda s, h, bt: backtest_method(s, forecast_mm, h, bt, window=6),
    "MM12": lambda s, h, bt: backtest_method(s, forecast_mm, h, bt, window=12),
    "NS":   lambda s, h, bt: backtest_method(s, forecast_naive_seasonal, h, bt),
    "CS":   lambda s, h, bt: backtest_method(s, forecast_seasonal_growth, h, bt),
    "TL":   lambda s, h, bt: backtest_method(s, forecast_trend_linear, h, bt),
    "ARIMA": lambda s, h, bt: backtest_method(
        s, lambda sr, h: forecast_arima(sr, h)["forecast"], h, bt),
    "ETS":  lambda s, h, bt: backtest_method(
        s, lambda sr, h: forecast_ets(sr, h)["forecast"], h, bt),
}


def run_all_forecasts(series: pd.Series, h: int, methods: list,
                      bt_window: int = 12, freq: str = "Mensuelle") -> dict:
    """
    Lance toutes les méthodes de prévision + backtesting pour une série.

    Parameters
    ----------
    freq : str, ignoré pour l'instant (réservé pour extension trimestrielle)

    Returns
    -------
    dict avec 'forecasts' (method→array), 'backtesting' (method→{mape,rmse}),
    'best_method' (str)
    """
    forecasts = {}
    backtesting = {}

    for m in methods:
        # Prévision
        result = METHOD_DISPATCH[m](series, h)
        if isinstance(result, dict):
            forecasts[m] = result["forecast"]
        else:
            forecasts[m] = result

        # Backtesting
        bt = BACKTEST_DISPATCH[m](series, h, bt_window)
        backtesting[m] = bt

    # Meilleure méthode par MAPE
    valid = {m: bt["mape"] for m, bt in backtesting.items()
             if not np.isnan(bt["mape"])}
    best = min(valid, key=valid.get) if valid else methods[0]

    return {"forecasts": forecasts, "backtesting": backtesting,
            "best_method": best}
