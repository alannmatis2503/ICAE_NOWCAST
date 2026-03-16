"""Page Documentation / Aide."""
import streamlit as st

st.title("📄 Documentation")
st.markdown("""
## Structure attendue du fichier ICAE Pays consolidé

```
📁 ICAE_XXX_Consolide.xlsx
├── 📄 Consignes
│   ├── Ligne 19 : Année de base
│   └── Ligne 20 : Plage de lignes base (ex: 124:135)
├── 📄 Codification
│   └── Code | Label | Unité/Source | Secteur | PRIOR | Statut
├── 📄 Donnees_calcul
│   ├── Colonne A : Dates (YYYY-MM-DD)
│   └── Colonnes B+ : Séries temporelles brutes
├── 📄 CALCUL_ICAE
│   ├── Ligne 1 : Inv_Ecart_Type (1/STDEV)
│   ├── Ligne 2 : Pondération (a)
│   ├── Ligne 9 : PRIOR MEDIAN (b)
│   ├── Ligne 12 : Pondération finale
│   ├── Lignes 14-15 : En-têtes
│   └── Lignes 16+ : TCS par variable
├── 📄 Contrib
│   └── Contributions sectorielles par mois
└── 📄 Resultats_Trim
    └── ICAE trimestriel + GA + GT + contributions
```

## Structure du fichier CEMAC

```
📁 ICAE_CEMAC_Consolide.xlsx
├── 📄 Poids_PIB → Code, Pays, PIB, Poids
├── 📄 ICAE_Pays → Date, ICAE par pays, ICAE CEMAC, GA, GT
├── 📄 ICAE_Trimestriel → Trimestre, ICAE par pays, GA Trim
└── 📄 Consignes → Mode d'emploi
```

## Formules mathématiques

### Taux de Croissance Symétrique (TCS)
$$C_{j,t} = 200 \\times \\frac{X_{j,t} - X_{j,t-1}}{X_{j,t} + X_{j,t-1}}$$

### Pondérations
$$\\omega_j = \\frac{1}{\\sigma_j}, \\quad a_j = \\frac{\\omega_j}{\\sum \\omega_j}, \\quad Pond\\_finale_j = \\frac{a_j \\times b_j}{\\sum (a_j \\times b_j)}$$

### Indice récursif
$$I_t = I_{t-1} \\times \\frac{200 + \\Sigma m_t}{200 - \\Sigma m_t}$$

### Normalisation base 100
$$ICAE_t = 100 \\times \\frac{I_t}{\\overline{I}_{\\text{année de base}}}$$

### ICAE CEMAC agrégé
$$ICAE\\_CEMAC_t = \\sum_{i=1}^{6} w_i \\times ICAE_{i,t}, \\quad w_i = \\frac{PIB_i}{\\sum PIB_i}$$

## Modules

| Module | Description |
|--------|-------------|
| **Module 1 — ICAE** | Calcul de l'ICAE à partir du fichier consolidé |
| **Module 2 — Prévisions** | Prolongation des séries par 8 méthodes (MM3/6/12, NS, CS, TL, ARIMA, ETS) |
| **Module 3 — Nowcast** | Estimation du PIB en temps réel (Bridge, U-MIDAS, PC, DFM) |
| **Module 4 — CEMAC** | Agrégation pondérée de l'ICAE des 6 pays |
| **Module 5 — Rapports** | Génération de Notes Word (ICAE + Nowcast) |

## Profils pays

| Pays | Code | Variables | Année base | Notes |
|------|------|-----------|------------|-------|
| Cameroun | CMR | 22 | 2023 | Standard |
| Congo | CNG | 19 | 2023 | Données démarrent mars 2014 |
| Gabon | GAB | 24 | 2023 | Standard |
| Guinée Éq. | GNQ | 18 | 2023 | Poids égaux (Inv_ET=1) |
| RCA | RCA | 22 | 2023 | Standard |
| Tchad | TCD | 9 | **2015** | Base différente ! |
""")
