"""Moteur de prévision — méthodes naïves + ARIMA/ETS (multi-fréquence)."""
import numpy as np
import pandas as pd
import warnings

# Saisonnalités par fréquence
SEASON_MAP = {"Mensuelle": 12, "Trimestrielle": 4, "Annuelle": 1}


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
    if season <= 1:
        return np.full(h, vals[-1] if len(vals) > 0 else np.nan)
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
    """Croissance saisonnière : X_{t+h} = X_{t+h-s} * X_t / X_{t-s}."""
    vals = series.dropna().values
    n = len(vals)
    if season <= 1:
        # Pour annuel : tendance simple par taux de croissance moyen
        if n < 2:
            return np.full(h, np.nan)
        growth = vals[-1] / vals[-2] if vals[-2] != 0 else 1.0
        return np.array([vals[-1] * growth ** (i + 1) for i in range(h)])
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
    """Tendance linéaire sur les `window` dernières observations."""
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
def forecast_arima(series: pd.Series, h: int, season: int = 12) -> dict:
    """Auto-ARIMA via pmdarima."""
    try:
        import pmdarima as pm
        vals = series.dropna().values
        min_obs = max(2 * season, 10) if season > 1 else 6
        if len(vals) < min_obs:
            return {"forecast": np.full(h, np.nan), "conf_int": None}
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            model = pm.auto_arima(
                vals, seasonal=(season > 1), m=max(season, 1),
                stepwise=True, suppress_warnings=True,
                error_action="ignore",
                max_order=5, max_p=3, max_q=3,
            )
            fc, conf = model.predict(n_periods=h, return_conf_int=True,
                                     alpha=0.2)
        return {"forecast": fc, "conf_int": conf}
    except Exception:
        return {"forecast": np.full(h, np.nan), "conf_int": None}


def forecast_ets(series: pd.Series, h: int, season: int = 12) -> dict:
    """ETS via statsmodels."""
    try:
        from statsmodels.tsa.holtwinters import ExponentialSmoothing
        vals = series.dropna().values
        min_obs = max(2 * season, 10) if season > 1 else 6
        if len(vals) < min_obs:
            return {"forecast": np.full(h, np.nan), "conf_int": None}
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            if season > 1:
                model = ExponentialSmoothing(
                    vals, seasonal_periods=season,
                    trend="add", seasonal="add",
                    initialization_method="estimated",
                ).fit(optimized=True)
            else:
                model = ExponentialSmoothing(
                    vals, trend="add", seasonal=None,
                    initialization_method="estimated",
                ).fit(optimized=True)
            fc = model.forecast(h)
            residuals = model.resid
            se = np.std(residuals) * np.sqrt(np.arange(1, h + 1))
            lo = fc - 1.28 * se
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
    min_train = max(kwargs.get("season", 12) * 2, 10) if kwargs.get("season", 12) > 1 else 6
    if n < bt_window + h + min_train:
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
# Dispatch — paramétré par la saisonnalité
# ────────────────────────────────────────────────────────────────────────────
def _build_dispatch(season: int):
    """Construit les dictionnaires de dispatch pour une saisonnalité donnée."""
    # Fenêtre tendance linéaire adaptée
    tl_window = {12: 24, 4: 8, 1: 5}.get(season, 24)

    # Fenêtres de moyenne mobile adaptées
    if season >= 12:
        mm_windows = {"MM3": 3, "MM6": 6, "MM12": 12}
    elif season >= 4:
        mm_windows = {"MM2": 2, "MM4": 4, "MM8": 8}
    else:
        mm_windows = {"MM2": 2, "MM3": 3, "MM5": 5}

    method_dispatch = {}
    backtest_dispatch = {}

    for name, w in mm_windows.items():
        method_dispatch[name] = lambda s, h, _w=w: forecast_mm(s, h, _w)
        backtest_dispatch[name] = lambda s, h, bt, _w=w: backtest_method(
            s, forecast_mm, h, bt, window=_w)

    method_dispatch["NS"] = lambda s, h: forecast_naive_seasonal(s, h, season)
    backtest_dispatch["NS"] = lambda s, h, bt: backtest_method(
        s, forecast_naive_seasonal, h, bt, season=season)

    method_dispatch["CS"] = lambda s, h: forecast_seasonal_growth(s, h, season)
    backtest_dispatch["CS"] = lambda s, h, bt: backtest_method(
        s, forecast_seasonal_growth, h, bt, season=season)

    method_dispatch["TL"] = lambda s, h: forecast_trend_linear(s, h, tl_window)
    backtest_dispatch["TL"] = lambda s, h, bt: backtest_method(
        s, forecast_trend_linear, h, bt, window=tl_window)

    method_dispatch["ARIMA"] = lambda s, h: forecast_arima(s, h, season)
    backtest_dispatch["ARIMA"] = lambda s, h, bt: backtest_method(
        s, lambda sr, h, season=season: forecast_arima(sr, h, season)["forecast"],
        h, bt)

    method_dispatch["ETS"] = lambda s, h: forecast_ets(s, h, season)
    backtest_dispatch["ETS"] = lambda s, h, bt: backtest_method(
        s, lambda sr, h, season=season: forecast_ets(sr, h, season)["forecast"],
        h, bt)

    return method_dispatch, backtest_dispatch


def get_methods_for_frequency(freq: str) -> list:
    """Retourne la liste des méthodes disponibles pour une fréquence."""
    season = SEASON_MAP.get(freq, 12)
    dispatch, _ = _build_dispatch(season)
    return list(dispatch.keys())


# Dispatch mensuel par défaut (rétrocompatibilité)
METHOD_DISPATCH, BACKTEST_DISPATCH = _build_dispatch(12)


def run_all_forecasts(series: pd.Series, h: int, methods: list,
                      bt_window: int = 12, freq: str = "Mensuelle") -> dict:
    """
    Lance toutes les méthodes de prévision + backtesting pour une série.

    Parameters
    ----------
    freq : "Mensuelle", "Trimestrielle" ou "Annuelle"

    Returns
    -------
    dict avec 'forecasts', 'backtesting', 'best_method'
    """
    season = SEASON_MAP.get(freq, 12)
    method_dispatch, backtest_dispatch = _build_dispatch(season)

    forecasts = {}
    backtesting = {}

    for m in methods:
        if m not in method_dispatch:
            continue
        result = method_dispatch[m](series, h)
        if isinstance(result, dict):
            forecasts[m] = result["forecast"]
        else:
            forecasts[m] = result

        bt = backtest_dispatch[m](series, h, bt_window)
        backtesting[m] = bt

    # Meilleure méthode par MAPE
    valid = {m: bt["mape"] for m, bt in backtesting.items()
             if not np.isnan(bt["mape"])}
    best = min(valid, key=valid.get) if valid else methods[0]

    return {"forecasts": forecasts, "backtesting": backtesting,
            "best_method": best}
