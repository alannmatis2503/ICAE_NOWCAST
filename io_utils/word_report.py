"""Génération de rapports Word (.docx) — Note ICAE et Note Nowcast."""
import io
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from pathlib import Path


def _add_heading_beac(doc, text, level=1):
    """Ajoute un titre formaté BEAC."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = "Times New Roman"


def _add_body_text(doc, text):
    """Ajoute un paragraphe de texte corps."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(13)
    run.font.name = "Times New Roman"
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def _add_figure_caption(doc, caption):
    """Ajoute une légende de figure."""
    p = doc.add_paragraph()
    run = p.add_run(caption)
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = "Times New Roman"
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _add_image(doc, image_bytes, width=Cm(15)):
    """Ajoute une image depuis des bytes."""
    buf = io.BytesIO(image_bytes)
    doc.add_picture(buf, width=width)
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _add_table(doc, df: pd.DataFrame):
    """Ajoute un DataFrame comme table Word."""
    table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # En-têtes
    for j, col in enumerate(df.columns):
        cell = table.cell(0, j)
        cell.text = str(col)
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(10)
                run.font.name = "Times New Roman"

    # Données
    for i, (_, row) in enumerate(df.iterrows()):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            if isinstance(val, float):
                cell.text = f"{val:.2f}" if not np.isnan(val) else ""
            else:
                cell.text = str(val) if val is not None else ""
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(10)
                    run.font.name = "Times New Roman"

    return table


# ────────────────────────────────────────────────────────────────────────────
# Note ICAE CEMAC (format court, calqué sur le modèle existant)
# ────────────────────────────────────────────────────────────────────────────
def generate_note_icae(
    country_name: str,
    trimestre: str,
    annee: int,
    icae_data: dict,
    chart_evolution_bytes: bytes = None,
    chart_perspectives_bytes: bytes = None,
    logo_path: str = None,
    user_texts: dict = None,
) -> bytes:
    """
    Génère la Note ICAE (CEMAC ou pays).

    Parameters
    ----------
    user_texts : dict avec clés optionnelles :
        'accroche_evolution', 'paragraphs_evolution' (list[str]),
        'accroche_perspectives', 'paragraphs_perspectives' (list[str])
    """
    doc = Document()

    # Style par défaut
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(13)

    if user_texts is None:
        user_texts = {}

    # ── Logo ──────────────────────────────────────────────────────────────
    if logo_path and Path(logo_path).exists():
        doc.add_picture(str(logo_path), width=Cm(4))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── Titre ─────────────────────────────────────────────────────────────
    title = (
        f"EVOLUTION DE L'INDICE COMPOSITE DES ACTIVITES ECONOMIQUES (ICAE) "
        f"DE {country_name.upper()} AU {trimestre} {annee} "
        f"ET PERSPECTIVES A COURT TERME"
    )
    p = doc.add_paragraph()
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(13)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()  # espace

    # ── Section 1 : Evolution récente ─────────────────────────────────────
    _add_heading_beac(doc, "Evolution récente")

    # Accroche
    accroche_evo = user_texts.get(
        "accroche_evolution",
        f"[Renseignez l'accroche : Le rythme de progression des activités "
        f"économiques a ... au {trimestre} {annee}.]"
    )
    p = doc.add_paragraph()
    run = p.add_run(accroche_evo)
    run.bold = True
    run.font.size = Pt(14)

    # Paragraphes d'analyse
    paras_evo = user_texts.get("paragraphs_evolution", [
        "[Paragraphe 1 : Dynamique globale du secteur productif. "
        "Utilisez les données et graphiques ci-contre pour rédiger.]",
        "[Paragraphe 2 : Activités extractives — pétrole, mines, etc.]",
        "[Paragraphe 3 : Industrie manufacturière, BTP, sylviculture.]",
        "[Paragraphe 4 : Transports, services et autres secteurs.]",
    ])
    for para in paras_evo:
        _add_body_text(doc, para)

    # Figure 1
    _add_figure_caption(
        doc,
        f"Figure 1 : Evolution de l'ICAE {country_name} en glissement annuel"
    )
    if chart_evolution_bytes:
        _add_image(doc, chart_evolution_bytes)
    else:
        _add_body_text(doc, "[Graphique non disponible — générez-le dans le module ICAE]")

    doc.add_paragraph()  # espace

    # ── Section 2 : Perspectives ──────────────────────────────────────────
    _add_heading_beac(doc, "Perspectives d'évolution de l'ICAE à court terme")

    accroche_persp = user_texts.get(
        "accroche_perspectives",
        f"[Renseignez l'accroche : ...le rythme de progression des activités "
        f"économiques devrait... au trimestre suivant.]"
    )
    p = doc.add_paragraph()
    run = p.add_run(accroche_persp)
    run.bold = True
    run.font.size = Pt(14)

    paras_persp = user_texts.get("paragraphs_perspectives", [
        "[Paragraphe 1 : Projections sectorielles.]",
        "[Paragraphe 2 : Facteurs de risque et incertitudes.]",
        "[Paragraphe 3 : Synthèse — chiffres clés de la projection.]",
    ])
    for para in paras_persp:
        _add_body_text(doc, para)

    # Figure 2
    _add_figure_caption(
        doc,
        f"Figure 2 : Evolution de l'ICAE {country_name} en glissement annuel "
        f"(avec projection)"
    )
    if chart_perspectives_bytes:
        _add_image(doc, chart_perspectives_bytes)
    else:
        _add_body_text(doc, "[Graphique non disponible — générez-le dans le module Prévisions]")

    # Sauvegarder en bytes
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ────────────────────────────────────────────────────────────────────────────
# Note Nowcast (format technique)
# ────────────────────────────────────────────────────────────────────────────
def generate_note_nowcast(
    results_by_country: dict,
    chart_bytes: dict = None,
    logo_path: str = None,
    params: dict = None,
) -> bytes:
    """Génère la Note Nowcast technique."""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    if chart_bytes is None:
        chart_bytes = {}

    # Logo
    if logo_path and Path(logo_path).exists():
        doc.add_picture(str(logo_path), width=Cm(4))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("NOTE SUR LE NOWCAST DU PIB", level=0)

    # 1. Introduction
    doc.add_heading("1. Introduction", level=1)
    _add_body_text(doc, "[Contexte et objectifs du nowcasting.]")

    # 2. Organisation des travaux
    doc.add_heading("2. Organisation des travaux", level=1)
    _add_body_text(doc, "[Sources de données et méthodologie.]")

    # 4. Méthodes
    doc.add_heading("3. Explication des méthodes", level=1)
    for method in ["DFM (Dynamic Factor Model)", "Bridge Equation",
                    "U-MIDAS", "Principal Components"]:
        doc.add_heading(f"3.x {method}", level=2)
        _add_body_text(doc, f"[Description de la méthode {method}.]")

    # 6. Application empirique
    doc.add_heading("4. Application empirique", level=1)
    for code, results in results_by_country.items():
        doc.add_heading(f"4.x {code}", level=2)

        # Table de performance
        if "metrics_df" in results:
            _add_table(doc, results["metrics_df"])

        # Graphique
        if code in chart_bytes:
            _add_image(doc, chart_bytes[code])

    # Conclusion
    doc.add_heading("5. Conclusion", level=1)
    _add_body_text(doc, "[Conclusion et recommandations.]")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()
