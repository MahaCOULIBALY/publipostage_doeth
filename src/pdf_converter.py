"""
Module de conversion DOCX → PDF via l'automation COM de Microsoft Word.

Stratégie :
    - Une seule instance Word est ouverte pour tout le batch (perf x10 vs instance/fichier).
    - Context manager garantit la fermeture propre même en cas d'erreur (RAII).
    - Requiert Microsoft Word installé sur le poste (contexte Windows métier DOETH).
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Format Word : wdFormatPDF = 17
_WD_FORMAT_PDF = 17


class WordPDFConverter:
    """
    Context manager wrappant une unique instance Word COM pour la conversion batch DOCX→PDF.

    Usage :
        with WordPDFConverter() as converter:
            pdf_path = converter.convert(docx_path)
    """

    def __init__(self) -> None:
        self._word = None

    def __enter__(self) -> "WordPDFConverter":
        try:
            import win32com.client  # type: ignore
            self._word = win32com.client.Dispatch("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = False
            logger.debug("Instance Word COM ouverte")
        except ImportError:
            raise RuntimeError(
                "pywin32 est requis pour la conversion PDF. "
                "Installez-le avec : pip install pywin32"
            )
        except Exception as e:
            raise RuntimeError(
                f"Impossible d'ouvrir Microsoft Word : {e}") from e
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._word is not None:
            try:
                self._word.Quit()
                logger.debug("Instance Word COM fermée")
            except Exception as e:
                logger.warning(f"Erreur lors de la fermeture de Word : {e}")
        # Ne pas supprimer l'exception — la laisser remonter
        return False

    def convert(self, docx_path: Path) -> Path:
        """
        Convertit un fichier DOCX en PDF.

        Args:
            docx_path: Chemin absolu vers le fichier .docx source.

        Returns:
            Chemin vers le fichier .pdf créé (même dossier, même nom).

        Raises:
            RuntimeError: Si Word n'est pas initialisé ou si la conversion échoue.
        """
        if self._word is None:
            raise RuntimeError(
                "WordPDFConverter doit être utilisé comme context manager.")

        docx_path = Path(docx_path).resolve()
        pdf_path = docx_path.with_suffix(".pdf")

        doc = None
        try:
            doc = self._word.Documents.Open(str(docx_path))
            doc.SaveAs(str(pdf_path), FileFormat=_WD_FORMAT_PDF)
            logger.debug(f"PDF généré : {pdf_path.name}")
            return pdf_path
        except Exception as e:
            raise RuntimeError(
                f"Échec conversion PDF pour {docx_path.name} : {e}") from e
        finally:
            if doc is not None:
                try:
                    doc.Close(False)  # False = sans sauvegarder les modifs
                except Exception:
                    pass


def convert_batch(
    docx_paths: list[Path],
    delete_docx: bool = False,
    logger: Optional[logging.Logger] = None,
) -> list[Path]:
    """
    Convertit une liste de fichiers DOCX en PDF avec une unique instance Word.

    Args:
        docx_paths:   Liste des chemins DOCX à convertir.
        delete_docx:  Si True, supprime le DOCX source après conversion réussie.
        logger:       Logger optionnel (utilise le logger du module sinon).

    Returns:
        Liste des chemins PDF créés avec succès.
    """
    log = logger or globals()["logger"]
    pdf_paths: list[Path] = []

    if not docx_paths:
        return pdf_paths

    log.info(f"Conversion PDF : {len(docx_paths)} fichier(s) à traiter")

    with WordPDFConverter() as converter:
        for docx_path in docx_paths:
            try:
                pdf_path = converter.convert(Path(docx_path))
                pdf_paths.append(pdf_path)

                if delete_docx:
                    Path(docx_path).unlink(missing_ok=True)
                    log.debug(f"DOCX supprimé : {Path(docx_path).name}")

            except Exception as e:
                log.error(f"Erreur conversion {Path(docx_path).name} : {e}")

    log.info(
        f"Conversion PDF terminée : {len(pdf_paths)}/{len(docx_paths)} succès")
    return pdf_paths
