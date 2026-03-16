"""Composants UI réutilisables pour Streamlit."""
import streamlit as st
import pandas as pd
from pathlib import Path
from io_utils.excel_reader import list_sheets


def file_uploader_with_sheet(label: str, key: str,
                             default_sheet: str = None,
                             file_types=None):
    """
    Upload de fichier Excel + sélection de feuille.
    Retourne (filepath_or_bytes, sheet_name) ou (None, None).
    """
    if file_types is None:
        file_types = ["xlsx", "xls"]

    uploaded = st.file_uploader(label, type=file_types, key=f"upload_{key}")
    if uploaded is None:
        return None, None

    try:
        sheets = list_sheets(uploaded)
        uploaded.seek(0)
    except Exception as e:
        st.error(f"Erreur à la lecture du fichier : {e}")
        return None, None

    if default_sheet and default_sheet in sheets:
        default_idx = sheets.index(default_sheet)
    else:
        default_idx = 0
        if default_sheet:
            st.warning(f"⚠️ Feuille '{default_sheet}' non trouvée. "
                       f"Veuillez sélectionner manuellement.")

    selected_sheet = st.selectbox(
        f"Feuille à utiliser ({key})", sheets,
        index=default_idx, key=f"sheet_{key}",
    )

    return uploaded, selected_sheet


def country_selector(key: str = "country"):
    """Sélecteur de pays."""
    from config import COUNTRY_CODES, COUNTRY_NAMES
    options = COUNTRY_CODES
    labels = [f"{c} — {COUNTRY_NAMES[c]}" for c in options]
    selected = st.selectbox("Pays", labels, key=key)
    idx = labels.index(selected)
    return options[idx]


def display_defaults_table(params: dict, title: str = "Paramètres détectés"):
    """Affiche un tableau des valeurs par défaut détectées."""
    st.subheader(title)
    df = pd.DataFrame([
        {"Paramètre": k, "Valeur": v}
        for k, v in params.items()
    ])
    st.dataframe(df, use_container_width=True, hide_index=True)


def download_button(data: bytes, filename: str, label: str = "📥 Télécharger",
                    mime: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    """Bouton de téléchargement standardisé."""
    st.download_button(
        label=label,
        data=data,
        file_name=filename,
        mime=mime,
    )


def color_forecast_df(df: pd.DataFrame, forecast_col: str = "type"):
    """Applique un style coloré aux lignes de prévision."""
    def _style(row):
        if forecast_col in row.index and row[forecast_col] == "forecast":
            return ["background-color: #FDE9D9; color: #E26B0A"] * len(row)
        return [""] * len(row)
    return df.style.apply(_style, axis=1)


def metric_color(value, thresholds=(5, 10, 20)):
    """Retourne une couleur selon le MAPE."""
    if pd.isna(value):
        return "gray"
    if value < thresholds[0]:
        return "green"
    elif value < thresholds[1]:
        return "orange"
    else:
        return "red"
