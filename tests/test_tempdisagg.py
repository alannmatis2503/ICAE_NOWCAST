"""Test du module tempdisagg — toutes les méthodes."""
import sys
import os
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import numpy as np
import pandas as pd

from core.tempdisagg import (
    chow_lin, denton_cholette, fernandez,
    disaggregate_annual, disaggregate_annual_to_quarterly,
    _build_C, _recalibrate
)

print("=== Test matrice C ===")
C = _build_C(8, 2, s=4)
assert C.shape == (2, 8)
assert C[0, 0:4].sum() == 4.0
assert C[1, 4:8].sum() == 4.0
print("OK matrice C")

print("\n=== Test Denton-Cholette ===")
y_a = np.array([100.0, 120.0])
y_hf = denton_cholette(y_a, 8, s=4)
assert len(y_hf) == 8
assert np.allclose(y_hf[:4].sum(), 100.0)
assert np.allclose(y_hf[4:8].sum(), 120.0)
print(f"OK Denton: {y_hf}")

print("\n=== Test Chow-Lin ===")
np.random.seed(42)
X_hf = np.column_stack([np.ones(8), np.random.randn(8)])
y_hf_cl = chow_lin(y_a, X_hf, s=4)
assert len(y_hf_cl) == 8
y_hf_cl = _recalibrate(y_hf_cl, y_a, 2, s=4)
assert np.allclose(y_hf_cl[:4].sum(), 100.0, atol=0.1)
assert np.allclose(y_hf_cl[4:8].sum(), 120.0, atol=0.1)
print(f"OK Chow-Lin: {y_hf_cl.round(2)}")

print("\n=== Test Fernandez (Ecotrim) ===")
y_hf_f = fernandez(y_a, X_hf, s=4)
assert len(y_hf_f) == 8
y_hf_f = _recalibrate(y_hf_f, y_a, 2, s=4)
assert np.allclose(y_hf_f[:4].sum(), 100.0, atol=0.1)
assert np.allclose(y_hf_f[4:8].sum(), 120.0, atol=0.1)
print(f"OK Fernandez: {y_hf_f.round(2)}")

print("\n=== Test Fernandez sans indicateur ===")
y_hf_f2 = fernandez(y_a, None, s=4)
y_hf_f2 = _recalibrate(y_hf_f2, y_a, 2, s=4)
assert np.allclose(y_hf_f2[:4].sum(), 100.0, atol=0.1)
print(f"OK Fernandez sans indicateur: {y_hf_f2.round(2)}")

print("\n=== Test pipeline disaggregate_annual ===")
pib = pd.Series([100.0, 120.0, 130.0], index=[2020, 2021, 2022])

# Méthode Denton
s_denton = disaggregate_annual(pib, "Trimestrielle", method="denton")
assert len(s_denton) == 12
assert isinstance(s_denton.index, pd.PeriodIndex)
assert np.allclose(s_denton[:4].sum(), 100.0)
print(f"OK pipeline denton: {len(s_denton)} trimestres")

# Méthode Ecotrim  
s_eco = disaggregate_annual(pib, "Trimestrielle", method="ecotrim")
assert len(s_eco) == 12
assert np.allclose(s_eco[:4].sum(), 100.0)
print(f"OK pipeline ecotrim: {len(s_eco)} trimestres")

# Alias rétrocompatible
s_compat = disaggregate_annual_to_quarterly(pib)
assert len(s_compat) == 12
print(f"OK alias rétrocompatible")

# Avec méthode explicite
s_compat2 = disaggregate_annual_to_quarterly(pib, method="ecotrim")
assert len(s_compat2) == 12
print(f"OK alias avec method=ecotrim")

print("\n=== TOUS LES TESTS TEMPDISAGG OK ===")
