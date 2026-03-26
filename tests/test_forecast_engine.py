"""Test du forecast_engine — vérifier MAE dans les résultats."""
import sys
import os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import numpy as np
import pandas as pd

from core.forecast_engine import BACKTEST_DISPATCH

# Créer une série test
dates = pd.date_range("2015-01-01", periods=120, freq="MS")
np.random.seed(42)
values = 100 + np.cumsum(np.random.randn(120) * 0.5)
series = pd.Series(values, index=dates)

print("=== Test backtest avec MAE ===")
for name in ["MM3", "MM6", "NS", "CS", "TL"]:
    result = BACKTEST_DISPATCH[name](series, 6, 12)
    assert "mape" in result, f"MAPE manquant pour {name}"
    assert "mae" in result, f"MAE manquant pour {name}"
    assert "rmse" in result, f"RMSE manquant pour {name}"
    print(f"OK {name}: MAPE={result['mape']:.2f}%, MAE={result['mae']:.2f}, RMSE={result['rmse']:.2f}")

print("\n=== TESTS FORECAST ENGINE OK ===")
