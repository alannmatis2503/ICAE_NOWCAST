"""Graphiques Plotly réutilisables."""
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from config import COLOR_HIST, COLOR_FCST, COLOR_HIST_BG, COLOR_FCST_BG


def chart_icae_monthly(dates, icae, title="ICAE mensuel",
                       fcst_start=None):
    """Graphique ICAE mensuel avec zone prévisions."""
    fig = go.Figure()
    dates = pd.to_datetime(dates)

    if fcst_start is not None:
        mask_hist = dates < fcst_start
        mask_fcst = dates >= fcst_start
        fig.add_trace(go.Scatter(
            x=dates[mask_hist], y=icae[mask_hist],
            mode="lines", name="Historique",
            line=dict(color=COLOR_HIST, width=2),
        ))
        fig.add_trace(go.Scatter(
            x=dates[mask_fcst], y=icae[mask_fcst],
            mode="lines", name="Prévision",
            line=dict(color=COLOR_FCST, width=2, dash="dash"),
        ))
    else:
        fig.add_trace(go.Scatter(
            x=dates, y=icae,
            mode="lines", name="ICAE",
            line=dict(color=COLOR_HIST, width=2),
        ))

    fig.update_layout(
        title=title,
        xaxis_title="Date", yaxis_title="ICAE",
        template="plotly_white",
        hovermode="x unified",
    )
    return fig


def chart_ga_bars(dates, ga, title="Glissement annuel (%)",
                  fcst_start=None):
    """Barres de glissement annuel."""
    fig = go.Figure()
    dates = pd.to_datetime(dates) if not isinstance(dates, pd.DatetimeIndex) else dates
    ga_arr = np.array(ga, dtype=float)

    colors = []
    for i, d in enumerate(dates):
        if fcst_start and d >= pd.Timestamp(fcst_start):
            colors.append(COLOR_FCST)
        elif ga_arr[i] >= 0:
            colors.append(COLOR_HIST)
        else:
            colors.append("#C00000")

    fig.add_trace(go.Bar(
        x=dates, y=ga_arr,
        marker_color=colors,
        name="GA (%)",
    ))
    fig.update_layout(
        title=title,
        xaxis_title="Date", yaxis_title="GA (%)",
        template="plotly_white",
    )
    return fig


def chart_contributions(dates, contrib_df, title="Contributions sectorielles"):
    """Barres empilées des contributions sectorielles."""
    fig = go.Figure()
    dates = pd.to_datetime(dates)
    colors = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
              "#5B9BD5", "#70AD47", "#264478", "#9B59B6"]

    cols = [c for c in contrib_df.columns if c != "Date"]
    for i, col in enumerate(cols):
        fig.add_trace(go.Bar(
            x=dates, y=contrib_df[col],
            name=col,
            marker_color=colors[i % len(colors)],
        ))

    fig.update_layout(
        barmode="relative",
        title=title,
        xaxis_title="Date", yaxis_title="Contribution",
        template="plotly_white",
    )
    return fig


def chart_nowcast(pib_q, results, title="PIB observé vs Nowcasts"):
    """PIB observé + estimations Nowcast."""
    fig = go.Figure()
    colors = {"Bridge": "#4472C4", "U-MIDAS": "#ED7D31",
              "PC": "#A5A5A5", "DFM": "#FFC000"}

    # PIB observé
    fig.add_trace(go.Scatter(
        x=pib_q.index.to_timestamp() if hasattr(pib_q.index, 'to_timestamp') else pib_q.index,
        y=pib_q.values,
        mode="lines+markers", name="PIB observé",
        line=dict(color="black", width=3),
    ))

    for name, r in results.items():
        fc = r["forecast"]
        idx = fc.index.to_timestamp() if hasattr(fc.index, 'to_timestamp') else fc.index
        fig.add_trace(go.Scatter(
            x=idx, y=fc.values,
            mode="lines", name=name,
            line=dict(color=colors.get(name, "#888"), width=2),
        ))

    fig.update_layout(
        title=title,
        xaxis_title="Trimestre", yaxis_title="PIB",
        template="plotly_white",
        hovermode="x unified",
    )
    return fig


def chart_forecast_comparison(series, forecasts, selected_method,
                              dates_hist, dates_fcst, var_name):
    """Graphique de comparaison des méthodes de prévision."""
    fig = go.Figure()
    method_colors = {
        "MM3": "#4472C4", "MM6": "#ED7D31", "MM12": "#A5A5A5",
        "NS": "#FFC000", "CS": "#5B9BD5", "TL": "#70AD47",
        "ARIMA": "#264478", "ETS": "#9B59B6",
    }

    # Historique
    fig.add_trace(go.Scatter(
        x=dates_hist, y=series,
        mode="lines", name="Historique",
        line=dict(color=COLOR_HIST, width=2),
    ))

    # Prévisions par méthode
    for method, values in forecasts.items():
        is_selected = method == selected_method
        fig.add_trace(go.Scatter(
            x=dates_fcst, y=values,
            mode="lines+markers", name=method,
            line=dict(
                color=method_colors.get(method, "#888"),
                width=4 if is_selected else 1,
                dash="solid" if is_selected else "dot",
            ),
            opacity=1.0 if is_selected else 0.5,
        ))

    fig.update_layout(
        title=f"Prévisions — {var_name}",
        xaxis_title="Date", yaxis_title="Valeur",
        template="plotly_white",
        hovermode="x unified",
    )
    return fig


def chart_ga_nowcast(pib_q, results, title="GA du PIB et des Nowcasts"):
    """Glissement annuel du PIB observé et des nowcasts."""
    fig = go.Figure()
    colors = {"PIB_observe": "black", "Bridge": "#4472C4",
              "U-MIDAS": "#ED7D31", "PC": "#A5A5A5", "DFM": "#FFC000"}

    # GA PIB observé
    ga_pib = (pib_q / pib_q.shift(4) - 1) * 100
    idx = ga_pib.index.to_timestamp() if hasattr(ga_pib.index, 'to_timestamp') else ga_pib.index
    fig.add_trace(go.Bar(
        x=idx, y=ga_pib.values,
        name="PIB observé",
        marker_color="black", opacity=0.6,
    ))

    for name, r in results.items():
        fc = r["forecast"]
        ga_fc = (fc / fc.shift(4) - 1) * 100
        fc_idx = ga_fc.index.to_timestamp() if hasattr(ga_fc.index, 'to_timestamp') else ga_fc.index
        fig.add_trace(go.Scatter(
            x=fc_idx, y=ga_fc.values,
            mode="lines+markers", name=f"GA {name}",
            line=dict(color=colors.get(name, "#888"), width=2),
        ))

    fig.update_layout(
        title=title,
        xaxis_title="Trimestre", yaxis_title="GA (%)",
        template="plotly_white",
    )
    return fig


def fig_to_png_bytes(fig, width=900, height=500) -> bytes:
    """Convertit un graphique Plotly en bytes PNG."""
    try:
        return fig.to_image(format="png", width=width, height=height,
                            engine="kaleido")
    except Exception:
        return None


# ────────────────────────────────────────────────────────────────────────────
# Graphique trimestriel : barres empilées contributions + courbe GA
# ────────────────────────────────────────────────────────────────────────────
SECTOR_COLORS = {
    "Produits de base": "#4472C4",
    "Demande privée": "#ED7D31",
    "Secteur public": "#A5A5A5",
    "Financement du secteur privé": "#FFC000",
}
_DEFAULT_COLORS = ["#5B9BD5", "#70AD47", "#264478", "#9B59B6",
                   "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000"]


def chart_quarterly_contrib_ga(
    contrib_trim: pd.DataFrame,
    ga_trim: pd.Series,
    trimestres: list[str],
    title: str = "Contributions sectorielles au GA de l'ICAE (trimestriel)",
    forecast_start_idx: int | None = None,
) -> go.Figure:
    """
    Barres empilées (contributions sectorielles) + courbe GA de l'ICAE,
    trimestriel.

    Parameters
    ----------
    contrib_trim : DataFrame avec colonnes = secteurs
    ga_trim : Series de GA trimestriel (%)
    trimestres : liste de labels trimestre (ex: ["2024T1", "2024T2", ...])
    forecast_start_idx : index à partir duquel les données sont prédites
                         (pointillés sur la courbe, hachures sur les barres)
    """
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    sector_cols = [c for c in contrib_trim.columns
                   if c not in ("trimestre", "Date", "_quarter")]

    # Barres empilées par secteur
    for i, sect in enumerate(sector_cols):
        vals = contrib_trim[sect].values
        color = SECTOR_COLORS.get(sect, _DEFAULT_COLORS[i % len(_DEFAULT_COLORS)])

        if forecast_start_idx is not None and forecast_start_idx < len(vals):
            # Partie historique
            hist_vals = list(vals[:forecast_start_idx]) + [None] * (len(vals) - forecast_start_idx)
            fig.add_trace(go.Bar(
                x=trimestres, y=hist_vals,
                name=sect, marker_color=color,
                legendgroup=sect, showlegend=True,
            ))
            # Partie prévue (pattern hachuré)
            fcst_vals = [None] * forecast_start_idx + list(vals[forecast_start_idx:])
            fig.add_trace(go.Bar(
                x=trimestres, y=fcst_vals,
                name=f"{sect} (proj.)", marker_color=color,
                marker_pattern_shape="/", opacity=0.6,
                legendgroup=sect, showlegend=False,
            ))
        else:
            fig.add_trace(go.Bar(
                x=trimestres, y=vals,
                name=sect, marker_color=color,
            ))

    # Courbe GA ICAE
    ga_vals = ga_trim.values if hasattr(ga_trim, "values") else ga_trim
    if forecast_start_idx is not None and forecast_start_idx < len(ga_vals):
        # Historique (trait plein)
        hist_ga = list(ga_vals[:forecast_start_idx]) + [ga_vals[forecast_start_idx]] if forecast_start_idx < len(ga_vals) else list(ga_vals[:forecast_start_idx])
        hist_x = trimestres[:forecast_start_idx + 1] if forecast_start_idx < len(trimestres) else trimestres[:forecast_start_idx]
        fig.add_trace(go.Scatter(
            x=hist_x, y=hist_ga,
            mode="lines+markers", name="GA ICAE (%)",
            line=dict(color="#C00000", width=3),
            marker=dict(size=8, color="#C00000"),
        ))
        # Prévision (pointillés)
        fcst_ga = [ga_vals[forecast_start_idx - 1]] + list(ga_vals[forecast_start_idx:]) if forecast_start_idx > 0 else list(ga_vals[forecast_start_idx:])
        fcst_x = trimestres[max(0, forecast_start_idx - 1):]
        fig.add_trace(go.Scatter(
            x=fcst_x, y=fcst_ga,
            mode="lines+markers", name="GA ICAE (proj.)",
            line=dict(color="#C00000", width=3, dash="dot"),
            marker=dict(size=8, color="#C00000", symbol="diamond"),
        ))
    else:
        fig.add_trace(go.Scatter(
            x=trimestres, y=ga_vals,
            mode="lines+markers", name="GA ICAE (%)",
            line=dict(color="#C00000", width=3),
            marker=dict(size=8, color="#C00000"),
        ))

    fig.update_layout(
        barmode="relative",
        title=dict(text=title, font=dict(size=14)),
        xaxis_title="Trimestre",
        yaxis_title="Contribution / GA (%)",
        template="plotly_white",
        hovermode="x unified",
        legend=dict(
            orientation="h", yanchor="bottom", y=-0.25,
            xanchor="center", x=0.5, font=dict(size=10),
        ),
        margin=dict(l=60, r=30, t=50, b=80),
    )
    return fig
