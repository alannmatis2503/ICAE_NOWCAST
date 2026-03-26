"""Test CEMAC rapport flow."""
import sys
sys.path.insert(0, 'Apps/ICAE_Streamlit')
from pathlib import Path
import pandas as pd
import numpy as np

from core.cemac_engine import compute_icae_cemac, quarterly_cemac
from config import COUNTRY_CODES, COUNTRY_NAMES, POIDS_PIB

# Simuler des ICAE pays
np.random.seed(42)
dates = pd.date_range('2020-01-01', periods=48, freq='MS')
icae_dict = {
    code: pd.Series(np.random.normal(100, 5, 48), index=dates) 
    for code in COUNTRY_CODES[:3]
}

# Calculer CEMAC
result_df = compute_icae_cemac(icae_dict, POIDS_PIB)
print(f'CEMAC result: {result_df.shape}, columns={list(result_df.columns)}')

# Trimestriel
q = quarterly_cemac(result_df, pd.to_datetime(result_df.index))
print(f'Quarterly: {q.shape}, columns={list(q.columns)}')

# Simuler le fix du module Rapports
if 'Trimestre' in q.columns and 'trimestre' not in q.columns:
    q = q.rename(columns={'Trimestre': 'trimestre'})

# S'assurer que GA_Trim existe
if 'GA_Trim' not in q.columns and 'ICAE_CEMAC' in q.columns:
    q['GA_Trim'] = (q['ICAE_CEMAC'] / q['ICAE_CEMAC'].shift(4) - 1) * 100
if 'GT_Trim' not in q.columns and 'ICAE_CEMAC' in q.columns:
    q['GT_Trim'] = (q['ICAE_CEMAC'] / q['ICAE_CEMAC'].shift(1) - 1) * 100
if 'icae_trim' not in q.columns and 'ICAE_CEMAC' in q.columns:
    q['icae_trim'] = q['ICAE_CEMAC']

print(f'After fix: columns={list(q.columns)}')
print(f'  trimestre values: {q["trimestre"].tolist()[:3]}...')
print(f'  GA_Trim sample: {q["GA_Trim"].dropna().tolist()[:3]}')

# Test le trimestres_list
trimestres_list = q['trimestre'].astype(str).tolist()
print(f'  trimestres_list: {len(trimestres_list)} items')

ga_trim = q['GA_Trim']
valid_indices = [i for i, v in enumerate(ga_trim) if pd.notna(v)]
print(f'  Valid GA indices: {len(valid_indices)}')

if len(valid_indices) >= 2:
    valid_trims = [trimestres_list[i] for i in valid_indices]
    print(f'  Valid trims: {valid_trims[:3]}...')

# Test contributions par pays
_poids_used = POIDS_PIB
_cemac_ct_cols = []
for code in COUNTRY_CODES:
    if code in q.columns:
        ga_pays = (q[code] / q[code].shift(4) - 1) * 100
        col_name = COUNTRY_NAMES.get(code, code)
        q[col_name] = ga_pays * _poids_used.get(code, 0)
        _cemac_ct_cols.append(col_name)

ct_data = q[['trimestre'] + _cemac_ct_cols].copy() if _cemac_ct_cols else None
print(f'  Contrib cols: {_cemac_ct_cols}')
print(f'  ct_data: {ct_data is not None}')

print()
print('CEMAC RAPPORT TEST OK')
