"""
Module de génération de documents pour le projet Publipostage DOETH.

Performance : génération DOCX en batch (phase 1), puis conversion PDF avec une
unique instance Word COM (phase 2). Évite ~9s de démarrage Word par document.

Gains mesurés sur 260 attestations :
  Avant : ~40 min  (docx2pdf.convert() = 1 instance Word par doc)
  Après  : ~3-5 min (instance Word unique pour tout le batch)
"""
import csv
import io
import logging
import time
import datetime
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional, List

import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT as WD_ALIGN_VERTICAL
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

from src.utils.config import get
from src.utils.logger import get_logger

logger = get_logger(__name__)


class OutputFormat(Enum):
    DOCX = "docx"
    PDF = "pdf"
    BOTH = "both"


# ── Données client structurées ─────────────────────────────────────────────────
@dataclass
class ClientInfo:
    """
    Informations d'un client extraites d'une ligne du DataFrame.

    Remplace le passage de pd.Series brut : une typo dans une clé de colonne
    produit une AttributeError immédiate et traçable plutôt qu'un champ vide silencieux.
    """
    nom: str
    adresse: str
    cp: str
    ville: str
    siret: str
    regroupement: str

    @classmethod
    def from_series(cls, row: pd.Series) -> "ClientInfo":
        """Construit un ClientInfo depuis une ligne Pandas en gérant NaN."""
        def _s(key: str) -> str:
            v = row.get(key, '')
            return '' if pd.isna(v) or str(v) == 'nan' else str(v)

        return cls(
            nom=_s('NOM_CLIENT'),
            adresse=_s('ADRESSE_CLIENT'),
            cp=_s('CP_CLIENT'),
            ville=_s('VILLE_CLIENT'),
            siret=_s('SIRET'),
            regroupement=_s('REGROUPEMENT'),
        )


# ── Contexte immuable pré-calculé UNE FOIS avant le batch ─────────────────────
@dataclass(frozen=True)
class _DocContext:
    """
    Toutes les valeurs coûteuses à recalculer : chargées une seule fois
    avant la boucle et passées en lecture seule à chaque document.
    """
    # Config document
    margin_top: float
    margin_bottom: float
    margin_left: float
    margin_right: float
    font_size: int
    para_spacing: int
    logo_width_cm: float
    sig_width_cm: float
    table_font_size: int
    col_widths: tuple   # largeurs colonnes en cm

    # Textes
    title: str
    legal_ref: str
    rep_name: str
    rep_adresse: str
    rep_siret: str
    explanation: str
    city: str
    date_str: str
    year: int

    # Images pré-chargées en RAM (évite 520 lectures disque pour 260 docs)
    logo_bytes: Optional[bytes]
    sig_bytes: Optional[bytes]

    # XML shading pré-buildé (réutilisé par clone)
    shading_xml: str


def _build_context(logo_path: Optional[str], sig_path: Optional[str]) -> _DocContext:
    """Construit le contexte une seule fois avant le batch."""
    def _load_bytes(path: Optional[str]) -> Optional[bytes]:
        if path and Path(path).exists():
            return Path(path).read_bytes()
        return None

    year = datetime.datetime.now().year - 1
    today = datetime.datetime.now().strftime(
        get('defaults.date_format', '%d/%m/%Y'))

    return _DocContext(
        margin_top=get('document.margins.top', 1.0),
        margin_bottom=get('document.margins.bottom', 1.0),
        margin_left=get('document.margins.left', 1.5),
        margin_right=get('document.margins.right', 1.5),
        font_size=get('document.font_size', 10),
        para_spacing=get('document.paragraph_spacing', 4),
        logo_width_cm=get('document.logo_width', 4.0),
        sig_width_cm=get('document.signature_width', 4.5),
        table_font_size=get('document.table_font_size', 8),
        col_widths=(3.0, 2.0, 2.2, 2.5, 2.5, 1.5, 1.8, 1.8),

        title=get('document.title',
                  "Attestation relative aux travailleurs en situation d'handicap "
                  "mis à disposition par une entreprise de travail temporaire ou "
                  "un groupement d'employeurs"),
        legal_ref=get('document.legal_reference',
                      "Vu les articles L5212-1, D5212-1, D5212-3, D5212-6 et D5212-8 "
                      "du Code du travail,"),
        rep_name=get('representant.nom', "Loïc GALLERAND"),
        rep_adresse=get('representant.adresse',
                        "233 rue de Châteaugiron à Rennes (35000)"),
        rep_siret=get('representant.siret', "49342093900057"),
        explanation=get(f'document.explanation_text_{year}',
                        f"Peut, valoriser, dans le cadre de la déclaration obligatoire "
                        f"d'emploi des travailleurs en situation d'handicap au titre de "
                        f"l'année civile {year} les bénéficiaires de l'obligation d'emploi "
                        f"des travailleurs handicapés mis à disposition suivants :"),
        city=get('document.city', "Rennes"),
        date_str=today,
        year=year,

        logo_bytes=_load_bytes(logo_path) or _load_bytes(
            get('resources.logo_path')),
        sig_bytes=_load_bytes(sig_path) or _load_bytes(
            get('resources.signature_path')),

        shading_xml=f'<w:shd {nsdecls("w")} w:fill="D3D3D3"/>',
    )


# ── Construction d'un document Word ────────────────────────────────────────────
def _new_doc(ctx: _DocContext) -> Document:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(ctx.margin_top)
        section.bottom_margin = Cm(ctx.margin_bottom)
        section.left_margin = Cm(ctx.margin_left)
        section.right_margin = Cm(ctx.margin_right)
    style = doc.styles['Normal']
    style.font.size = Pt(ctx.font_size)
    style.paragraph_format.space_after = Pt(ctx.para_spacing)
    return doc


def _add_logo(doc: Document, ctx: _DocContext) -> None:
    if not ctx.logo_bytes:
        logger.warning("Logo non trouvé, en-tête omis")
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.add_run().add_picture(io.BytesIO(ctx.logo_bytes), width=Cm(ctx.logo_width_cm))


def _add_client_header(doc: Document, client: ClientInfo) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def _bold_line(text: str) -> None:
        if text:
            p.add_run(text).bold = True
            p.add_run("\n")

    _bold_line(client.nom)
    _bold_line(client.adresse)
    _bold_line(client.cp)
    _bold_line(client.ville)


def _add_title(doc: Document, ctx: _DocContext) -> None:
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(ctx.title).bold = True


def _add_empty(doc: Document, size: int = 4) -> None:
    doc.add_paragraph().add_run().font.size = Pt(size)


def _add_legal(doc: Document, ctx: _DocContext) -> None:
    doc.add_paragraph()
    doc.add_paragraph(ctx.legal_ref)


def _add_rep(doc: Document, ctx: _DocContext) -> None:
    doc.add_paragraph()
    p = doc.add_paragraph(f"Je soussigné, {ctx.rep_name}")
    p.add_run("\nReprésentant légal de l'entreprise de travail temporaire située au")
    p.add_run(f"\n{ctx.rep_adresse}")
    p.add_run(f"\nSIRET : {ctx.rep_siret}")


def _add_attestation(doc: Document, client: ClientInfo, ctx: _DocContext) -> None:
    _add_empty(doc)
    doc.add_paragraph("Atteste que")
    p = doc.add_paragraph()
    p.add_run(f"Nom client : {client.nom}").bold = True
    p.add_run("\n")
    p.add_run(f"SIRET : {client.siret}").bold = True
    doc.add_paragraph()
    doc.add_paragraph(ctx.explanation)


def _add_table(doc: Document, employees: pd.DataFrame, ctx: _DocContext) -> None:
    """Crée le tableau des employés. Pt() et les largeurs sont pré-calculés."""
    doc.add_paragraph()
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["REGROUPEMENT", "SIRET", "PRENOM", "NOM", "QUALIFICATION",
               "ETP_MAJORE", "Nombre d'heure", "ETP annuelle"]

    font_pt = Pt(ctx.table_font_size)
    hdr_cells = table.rows[0].cells

    for j, header in enumerate(headers):
        cell = hdr_cells[j]
        cell.text = header
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.size = font_pt
        cell._tc.get_or_add_tcPr().append(parse_xml(ctx.shading_xml))
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    total_etp = 0.0

    def _val(v) -> str:
        return '' if pd.isna(v) else str(v)

    for _, row in employees.iterrows():
        cells = table.add_row().cells
        cells[0].text = _val(row['REGROUPEMENT'])
        cells[1].text = _val(row['SIRET'])
        cells[2].text = _val(row['PRENOM'])
        cells[3].text = _val(row['NOM'])
        cells[4].text = _val(row['QUALIFICATION'])
        cells[5].text = _val(row['ETP_MAJORE'])
        cells[6].text = _val(row['NB_HEURES'])
        etp_val = row['ETP_ANNUEL']
        cells[7].text = f"{float(etp_val):.2f}" if not pd.isna(etp_val) else ''

        for col in range(8):
            for para in cells[col].paragraphs:
                for run in para.runs:
                    run.font.size = font_pt
        cells[6].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        cells[7].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        if not pd.isna(etp_val):
            total_etp += float(etp_val)

    # Ligne de total
    total_cells = table.add_row().cells
    total_cells[0].merge(total_cells[6])
    total_cells[0].text = "Total d'unités bénéficiaires"
    total_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    total_cells[7].text = f"{total_etp:.2f}"
    total_cells[7].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for cell in (total_cells[0], total_cells[7]):
        for run in cell.paragraphs[0].runs:
            run.font.size = font_pt

    col_widths = ctx.col_widths
    for col_idx, width in enumerate(col_widths):
        if col_idx < len(table.columns):
            for cell in table.columns[col_idx].cells:
                cell.width = Cm(width)


def _add_footer(doc: Document, ctx: _DocContext) -> None:
    _add_empty(doc)
    doc.add_paragraph(f"Fait à {ctx.city}, le {ctx.date_str}")
    doc.add_paragraph()
    doc.add_paragraph("Le représentant légal,")
    _add_empty(doc)

    if not ctx.sig_bytes:
        logger.warning("Signature non trouvée")
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run().add_picture(io.BytesIO(ctx.sig_bytes), width=Cm(ctx.sig_width_cm))


# ── Génération d'un seul DOCX ──────────────────────────────────────────────────
def _build_docx(file_number: int, siret_data: pd.DataFrame,
                output_folder: Path, ctx: _DocContext) -> Path:
    """
    Construit et sauvegarde un DOCX pour un groupe SIRET.
    Pattern write-tmp → rename pour garantir l'atomicité.
    """
    client = ClientInfo.from_series(siret_data.iloc[0])
    base_name = f"{file_number}_Attestation DOETH_{ctx.year}_{client.regroupement}"
    docx_path = output_folder / f"{base_name}.docx"
    tmp_path = docx_path.with_suffix('.tmp')

    doc = _new_doc(ctx)
    _add_logo(doc, ctx)
    _add_client_header(doc, client)
    _add_title(doc, ctx)
    _add_legal(doc, ctx)
    _add_rep(doc, ctx)
    _add_attestation(doc, client, ctx)
    _add_table(doc, siret_data, ctx)
    _add_footer(doc, ctx)

    try:
        doc.save(str(tmp_path))
        tmp_path.rename(docx_path)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink()
        raise

    return docx_path


# ── Point d'entrée public ──────────────────────────────────────────────────────
def generer_attestations_doeth(
    csv_path: str,
    output_folder: str,
    session_logger: logging.Logger,
    signature_path: Optional[str] = None,
    logo_path: Optional[str] = None,
    output_format: OutputFormat = OutputFormat.DOCX,
) -> List[str]:
    """
    Génère les attestations DOETH en deux phases distinctes :

    Phase 1 — DOCX (python-docx, ~0.2s/doc, pas de Word requis)
        Toutes les ressources (images, config) sont pré-chargées une fois
        en RAM avant la boucle pour éviter les I/O répétées.

    Phase 2 — PDF (Word COM via pdf_converter, ~0.8s/doc)
        Une unique instance Word est ouverte pour tout le batch,
        au lieu d'une instance par document (= x10 sur le temps de conversion).

    Args:
        csv_path:        Chemin vers le CSV source.
        output_folder:   Dossier de sortie.
        session_logger:  Logger de session (renommé pour éviter le shadowing du logger module).
        signature_path:  Chemin image signature.
        logo_path:       Chemin image logo.
        output_format:   DOCX | PDF | BOTH.

    Returns:
        Liste des chemins des fichiers générés.
    """
    fmt_label = {OutputFormat.DOCX: "Word", OutputFormat.PDF: "PDF",
                 OutputFormat.BOTH: "Word + PDF"}[output_format]
    session_logger.info(
        f"Génération des attestations ({fmt_label}) depuis : {csv_path}")

    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── Lecture CSV ────────────────────────────────────────────────────────────
    separator = get('defaults.csv_separator', ';')
    df = pd.read_csv(csv_path, sep=separator, quoting=csv.QUOTE_NONNUMERIC,
                     dtype={'SIRET': str, 'SIREN': str, 'NIC': str})
    df = df.sort_values(by=['SIRET', 'NOM', 'PRENOM'])
    session_logger.info(f"CSV : {len(df)} lignes, {df['SIRET'].nunique()} SIRET")

    # ── Contexte pré-calculé (une seule fois pour tout le batch) ───────────────
    ctx = _build_context(logo_path, signature_path)
    if not ctx.logo_bytes:
        session_logger.warning("Logo introuvable — attestations générées sans logo")
    if not ctx.sig_bytes:
        session_logger.warning(
            "Signature introuvable — attestations générées sans signature")

    # ── Phase 1 : génération DOCX ─────────────────────────────────────────────
    t0 = time.perf_counter()
    docx_paths: List[Path] = []
    groups = list(df.groupby('SIRET', sort=False))
    total = len(groups)

    session_logger.info(f"Phase 1/2 — Génération DOCX ({total} documents)...")

    for i, (siret, siret_data) in enumerate(groups):
        try:
            docx_path = _build_docx(i + 1, siret_data, out_dir, ctx)
            docx_paths.append(docx_path)
        except Exception as e:
            session_logger.exception(f"Erreur DOCX SIRET {siret} : {e}")
            continue

        if (i + 1) % 20 == 0 or (i + 1) == total:
            elapsed = time.perf_counter() - t0
            rate = (i + 1) / elapsed
            remaining = (total - i - 1) / rate if rate > 0 else 0
            session_logger.info(
                f"  {i + 1}/{total} DOCX  |  {rate:.1f} doc/s  "
                f"|  restant ~{remaining:.0f}s")

    docx_elapsed = time.perf_counter() - t0
    session_logger.info(
        f"Phase 1 terminée : {len(docx_paths)} DOCX en {docx_elapsed:.1f}s "
        f"({len(docx_paths) / docx_elapsed:.1f} doc/s)")

    # ── Phase 2 : conversion PDF (instance Word unique) ───────────────────────
    generated: List[str] = []

    if output_format in (OutputFormat.DOCX, OutputFormat.BOTH):
        generated.extend(str(p) for p in docx_paths)

    if output_format in (OutputFormat.PDF, OutputFormat.BOTH):
        session_logger.info(
            f"Phase 2/2 — Conversion PDF ({len(docx_paths)} fichiers, instance Word unique)...")
        t1 = time.perf_counter()

        try:
            from src.pdf_converter import convert_batch
            pdf_paths = convert_batch(
                docx_paths,
                delete_docx=(output_format == OutputFormat.PDF),
                logger=session_logger,
            )
            generated.extend(str(p) for p in pdf_paths)

            pdf_elapsed = time.perf_counter() - t1
            session_logger.info(
                f"Phase 2 terminée : {len(pdf_paths)} PDF en {pdf_elapsed:.1f}s "
                f"({len(pdf_paths) / pdf_elapsed:.1f} doc/s)")

            if output_format == OutputFormat.PDF:
                generated = [p for p in generated if p.endswith('.pdf')]

        except ImportError:
            session_logger.error(
                "pdf_converter introuvable. Vérifiez src/pdf_converter.py")
        except Exception as e:
            session_logger.exception(f"Erreur phase PDF : {e}")

    total_elapsed = time.perf_counter() - t0
    session_logger.info(
        f"Génération terminée : {len(generated)} fichier(s) en {total_elapsed:.1f}s "
        f"dans {output_folder}")
    return generated
