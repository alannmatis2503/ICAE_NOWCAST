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

---

## Détection automatique des scénarios conjoncturels

Le module Rapports détecte automatiquement le scénario conjoncturel de chaque pays en
analysant l'évolution du **Glissement Annuel (GA) trimestriel** de l'ICAE. Voici les 7
scénarios implémentés :

| Scénario | Conditions | Exemple |
|----------|-----------|---------|
| **Reprise** | GA actuel > 0 **et** GA trimestre précédent < 0 | GA passe de −1,2 % à +2,5 % → l'activité renoue avec la croissance |
| **Retournement** | GA actuel < 0 **et** GA trimestre précédent > 0 | GA passe de +3,1 % à −0,8 % → bascule en zone de contraction |
| **Contraction aggravée** | GA actuel < 0, GA précédent < 0 **et** GA actuel < GA précédent | GA passe de −1,5 % à −3,2 % → le repli s'approfondit |
| **Contraction** | GA actuel < 0 (autres cas) | GA = −2,0 %, GA précédent = −3,5 % → l'activité reste en contraction mais se redresse |
| **Accélération** | GA actuel > GA précédent > 0 **et** GA actuel > GA il y a un an | GA passe de +2,0 % à +4,5 %, supérieur à celui d'il y a un an → dynamique haussière renforcée |
| **Ralentissement** | GA actuel > 0 **et** GA actuel < GA précédent | GA passe de +5,0 % à +3,2 % → l'activité ralentit |
| **Croissance stable** | GA actuel > 0 (autres cas) | GA = +3,0 %, GA précédent ≈ +3,1 % → rythme stable |

Chaque scénario génère automatiquement un texte d'accroche et un paragraphe d'analyse
adaptés, incluant les valeurs exactes du GA et les secteurs moteurs/freineurs.

---

## Sélection des variables pour le Nowcast (Module 3)

### Phase de sélection structurelle

Cette phase, à réaliser tous les **2-3 ans**, détermine quels indicateurs haute
fréquence sont les plus pertinents pour estimer le PIB en temps réel.

**1. Depuis la feuille Codification** (méthode par défaut) :
- Les variables avec `PRIOR > 0` et `Statut ≠ Inactif` sont automatiquement retenues
- Nombre typique : 9 à 24 variables par pays

**2. ACP + Corrélation avec le PIB** :
- Les données **mensuelles** sont **annualisées** avant l'ACP :
  - **Flux** (production, exportations, consommation…) → **somme** annuelle
  - **Taux** (IPC, ratio de créances…) → **moyenne** annuelle
  - **Stocks** (masse monétaire, crédit…) → **valeur de fin de période** (last)
- L'ACP identifie les composantes principales et les loadings
- Le **PIB** est traité comme **variable supplémentaire** : il est projeté sur le cercle
  de corrélation sans influencer le calcul des axes
- Les variables les plus proches du PIB sur le cercle (forte corrélation avec les axes
  sur lesquels le PIB se projette) sont les meilleures candidates

### Types d'agrégation mensuel → trimestriel

La classification par défaut utilise des mots-clés dans le nom de la variable :

| Type | Mots-clés reconnus | Agrégation |
|------|-------------------|------------|
| **Flux** | production, exportation, importation, chiffre d'affaires, recettes, dépenses, transport | Somme des 3 mois |
| **Taux** | taux, IPC, ratio, variation M2, créances douteuses | Moyenne des 3 mois |
| **Stock** | monnaie, M2, circulation fiduciaire, crédit à l'économie, créances nettes, inverse créances douteuses | Valeur du dernier mois |

L'utilisateur peut modifier manuellement le type d'agrégation pour chaque variable.

---

## Méthodes de trimestrialisation du PIB (Module 3 — Nowcast)

Le module Nowcast nécessite un PIB **trimestriel**. Si seul un PIB annuel est disponible,
l'application propose deux méthodes de désagrégation temporelle (trimestrialisation) :

---

### 1. Chow-Lin

**Principe :** Méthode par régression sur un indicateur haute fréquence (HF). Elle redistribue
le total annuel en trimestres en s'appuyant sur la dynamique d'un indicateur observé à
fréquence trimestrielle (ex : ICAE, indice de production).

**Modèle :**
$$Y_t = \\beta X_t + u_t, \\quad u_t = \\rho u_{t-1} + \\varepsilon_t$$

où $Y_t$ est le PIB trimestriel inconnu, $X_t$ l'indicateur HF, et $u_t$ un résidu autocorrélé
d'ordre 1 (AR1, coefficient $\\rho$ estimé par MV).

**Étapes :**
1. Régresser le PIB annuel $Y^A$ sur l'indicateur annualisé $X^A$ : $\\hat{\\beta} = (X^{A\\prime} \\Sigma^{-1} X^A)^{-1} X^{A\\prime} \\Sigma^{-1} Y^A$
2. Calculer les résidus annuels : $\\hat{u}^A = Y^A - X^A \\hat{\\beta}$
3. Distribuer les résidus aux trimestres via la matrice de dispersion de Chow-Lin
4. Résultat : $\\hat{Y}_t = X_t \\hat{\\beta} + D_t \\hat{u}^A$, avec contrainte $\\sum_{t \\in A} \\hat{Y}_t = Y^A$

**Usage :** Préférable quand un bon indicateur trimestriel est disponible et corrélé avec le PIB.

---

### 2. Denton-Cholette

**Principe :** Méthode de benchmarking par minimisation des variations. Elle répartit
le total annuel en trimestres en perturbant le moins possible le profil d'un indicateur
de référence (ou une interpolation linéaire si aucun indicateur n'est fourni).

**Fonction objectif (variante proportionnelle) :**
$$\\min_{\\hat{Y}} \\sum_{t=2}^{T} \\left(\\frac{\\hat{Y}_t}{\\hat{X}_t} - \\frac{\\hat{Y}_{t-1}}{\\hat{X}_{t-1}}\\right)^2
\\quad \\text{sous contrainte} \\quad \\sum_{t \\in A} \\hat{Y}_t = Y^A$$

où $\\hat{X}_t$ est l'indicateur (ou 1 si aucun indicateur), et la contrainte impose la
cohérence avec les totaux annuels.

**Variante additive :**
$$\\min_{\\hat{Y}} \\sum_{t=2}^{T} \\left[(\\hat{Y}_t - \\hat{X}_t) - (\\hat{Y}_{t-1} - \\hat{X}_{t-1})\\right]^2$$

**Usage :** Recommandé quand aucun indicateur HF n'est disponible ou en cas de doute
sur la qualité de l'indicateur.

---

### Contrainte d'agrégation temporelle

Dans les deux méthodes, la contrainte fondamentale est :
$$\\sum_{q=1}^{4} \\hat{Y}_{A,q} = Y^A \\quad \\forall A$$

(la somme des 4 trimestres reconstitue exactement le total annuel observé).

---

### Choix automatique dans l'application

| Condition | Méthode utilisée |
|-----------|-----------------|
| Indicateur HF fourni | Chow-Lin |
| Pas d'indicateur HF | Denton-Cholette |

---

## Méthodes de prévision (Module 2)

Le module de prévisions propose **8 méthodes** pour prolonger chaque série temporelle.
L'application les évalue toutes et recommande automatiquement la meilleure.

### Méthodes naïves

| Méthode | Description |
|---------|-------------|
| **MM3** | **Moyenne mobile 3** — Moyenne des 3 dernières observations, reproduite sur l'horizon. Simple et stable pour les séries peu volatiles. |
| **MM6** | **Moyenne mobile 6** — Idem avec une fenêtre de 6 mois. Lisse davantage les fluctuations. |
| **MM12** | **Moyenne mobile 12** — Fenêtre d'un an. Capture le niveau moyen annuel et élimine la saisonnalité. |
| **NS** | **Naïve Saisonnière** — $X_{t+h} = X_{t+h-12}$ : reproduit la valeur de la même période l'année précédente. Adaptée aux séries à profil saisonnier marqué. |
| **CS** | **Croissance Saisonnière** — Applique le taux de croissance annuel récent au profil saisonnier : $X_{t+h} = X_{t+h-12} \\times \\frac{X_t}{X_{t-12}}$. Capte à la fois la saisonnalité et la tendance. |
| **TL** | **Tendance Linéaire** — Régression linéaire sur les 24 dernières observations, puis extrapolation. Adaptée aux séries avec tendance régulière sans rupture. |

### Méthodes statistiques avancées

| Méthode | Description |
|---------|-------------|
| **ARIMA** | **Auto-ARIMA** — Modèle ARIMA saisonnier (m=12) sélectionné automatiquement par le critère AIC (via `pmdarima`). Identifie l'ordre optimal (p,d,q)(P,D,Q). Requiert au moins 24 observations. Fournit des intervalles de confiance à 80 %. |
| **ETS** | **Lissage exponentiel** — Holt-Winters (tendance + saisonnalité additive) via `statsmodels`. Décompose la série en niveau, tendance et composante saisonnière. Requiert au moins 24 observations. |

### Sélection automatique de la meilleure méthode

L'application recommande la meilleure méthode **par variable** en se basant sur le
**MAPE** (Mean Absolute Percentage Error) calculé par **backtesting glissant** :

1. **Validation glissante** : sur les 12 dernières observations, à chaque pas $t$,
   le modèle est entraîné sur les données jusqu'à $t-1$ et prédit la valeur en $t$.
2. **Calcul du MAPE** : $\\text{MAPE} = \\frac{100}{n} \\sum \\left| \\frac{A_t - P_t}{A_t} \\right|$
   où $A_t$ est la valeur observée et $P_t$ la prévision.
3. **Sélection** : la méthode avec le **MAPE le plus bas** est recommandée.

Trois métriques sont calculées pour chaque méthode : **MAPE**, **RMSE** et **MAE**.
L'utilisateur peut consulter le tableau comparatif et choisir un autre critère ou
remplacer manuellement la méthode recommandée.

---

## Modules

| Module | Description |
|--------|-------------|
| **Module 1 — ICAE** | Calcul de l'ICAE à partir du fichier consolidé |
| **Module 2 — Prévisions** | Prolongation des séries par 8 méthodes (MM3/6/12, NS, CS, TL, ARIMA, ETS) |
| **Module 3 — Nowcast** | Estimation du PIB en temps réel (Bridge, U-MIDAS, PC, DFM) |
| **Module 4 — Agrégation CEMAC** | Agrégation pondérée de l'ICAE et du PIB Nowcast des 6 pays |
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
