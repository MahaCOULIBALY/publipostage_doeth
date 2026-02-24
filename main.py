#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Point d'entrée principal du projet Publipostage DOETH.

Ce script orchestre l'ensemble du processus de traitement des données
et de génération des attestations DOETH.
"""

import csv
import sys
import time
import argparse
import datetime
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple, Callable

import pandas as pd

from src.utils.config import Config, get, init_config
from src.utils.logger import setup_logger, get_logger
from src.data_processor import nettoyer_fichier_excel
from src.document_generator import generer_attestations_doeth, OutputFormat


@dataclass
class RunStats:
    """Statistiques structurées d'une exécution — remplace le Dict[str, Any] non typé."""
    total_rows: int = 0
    unique_sirets: int = 0
    unique_clients: int = 0
    total_docs: int = 0
    total_etp: float = 0.0
    avg_etp_per_employee: float = 0.0
    avg_etp_per_siret: float = 0.0
    total_heures: float = 0.0
    avg_heures_per_employee: float = 0.0
    duration_seconds: float = 0.0
    regroupements: dict = field(default_factory=dict)
    files_by_format: dict = field(default_factory=dict)


def parse_arguments() -> argparse.Namespace:
    """Parse les arguments de ligne de commande."""
    parser = argparse.ArgumentParser(
        description="Publipostage DOETH - Génération d'attestations à partir de fichiers Excel"
    )

    parser.add_argument(
        "--config", type=str,
        help="Chemin vers un fichier de configuration personnalisé"
    )
    parser.add_argument(
        "--input", type=str,
        help="Chemin vers le fichier Excel d'entrée (override config)"
    )
    parser.add_argument(
        "--sheet", type=str,
        help="Nom de la feuille Excel à traiter (override config)"
    )
    parser.add_argument(
        "--output-dir", type=str,
        help="Dossier de sortie pour les attestations (override config)"
    )
    parser.add_argument(
        "--skip-processing", action="store_true",
        help="Ignorer l'étape de traitement de l'Excel et utiliser le CSV déjà généré"
    )
    parser.add_argument(
        "--csv-path", type=str,
        help="Chemin vers un fichier CSV déjà traité (utilisé avec --skip-processing)"
    )
    parser.add_argument(
        "--debug", action="store_true",
        help="Active le mode debug avec logs détaillés"
    )
    parser.add_argument(
        "--format", choices=["docx", "pdf", "both"], default="docx",
        help="Format de sortie des attestations : docx (défaut), pdf, both"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Simule le traitement complet sans écrire de fichiers Word/PDF"
    )

    return parser.parse_args()


def setup_environment(args: argparse.Namespace) -> Tuple[Dict[str, Any], logging.Logger]:
    """
    Configure l'environnement d'exécution et initialise les ressources.

    Crée UN SEUL logger pour toute la session et le retourne avec les paramètres,
    éliminant la double instanciation du logger précédente.

    Args:
        args: Arguments de ligne de commande (argparse.Namespace ou duck-type compatible)

    Returns:
        Tuple (params, logger) — paramètres d'exécution et logger de session
    """
    init_config(getattr(args, 'config', None))

    console_level = logging.DEBUG if args.debug else logging.INFO
    logs_dir = get('paths.logs_dir')
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')

    logger = setup_logger(
        logs_dir=logs_dir,
        name=f"publipostage_{timestamp}",
        console_level=console_level,
        file_level=logging.DEBUG,
        enable_colors=True,
    )

    logger.info("=== DÉMARRAGE DU PUBLIPOSTAGE DOETH ===")
    logger.info(f"Date et heure: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    params: Dict[str, Any] = {
        "timestamp": timestamp,
        "log_level": console_level,
        "logs_dir": logs_dir,
        "dry_run": getattr(args, 'dry_run', False),
    }

    # Chemins d'entrée/sortie
    if args.input:
        params["input_file"] = args.input
    else:
        input_dir = Path(get('paths.input_dir'))
        params["input_file"] = str(input_dir / get('defaults.input_filename'))

    params["sheet_name"] = args.sheet if args.sheet else get('defaults.excel_sheet')
    params["output_dir"] = args.output_dir if args.output_dir else get('paths.output_dir')
    params["skip_processing"] = args.skip_processing

    if args.csv_path:
        params["csv_path"] = args.csv_path
    else:
        processed_dir = Path(get('paths.processed_dir'))
        params["csv_path"] = str(processed_dir / f"processed_{timestamp}.csv")

    # Création des dossiers nécessaires
    for dir_path in [
        logs_dir,
        get('paths.input_dir'),
        get('paths.processed_dir'),
        get('paths.output_dir'),
    ]:
        Path(dir_path).mkdir(parents=True, exist_ok=True)

    # Vérification des ressources
    params["logo_path"] = get('resources.logo_path')
    params["signature_path"] = get('resources.signature_path')

    for name, path in [("Logo", params["logo_path"]), ("Signature", params["signature_path"])]:
        if not path or not Path(path).exists():
            logger.warning(f"{name} non trouvé: {path}")
        else:
            logger.info(f"{name} trouvé: {path}")

    # Validation du fichier d'entrée
    if not params["skip_processing"] and not Path(params["input_file"]).exists():
        logger.error(f"Fichier d'entrée non trouvé: {params['input_file']}")
        raise FileNotFoundError(f"Fichier d'entrée non trouvé: {params['input_file']}")

    logger.info("Environnement d'exécution configuré avec succès")
    return params, logger


def process_data(
    params: Dict[str, Any],
    logger: logging.Logger,
) -> Tuple[str, pd.DataFrame]:
    """
    Traite les données Excel pour préparer le fichier CSV intermédiaire.

    Returns:
        Tuple (csv_path, df) — le chemin du CSV ET le DataFrame en mémoire,
        ce qui évite un rechargement disque inutile dans main().
    """
    if params["skip_processing"]:
        csv_path = params["csv_path"]
        if not Path(csv_path).exists():
            logger.error(f"Fichier CSV spécifié non trouvé: {csv_path}")
            raise FileNotFoundError(f"Fichier CSV spécifié non trouvé: {csv_path}")

        logger.info(f"Étape de traitement ignorée, utilisation du CSV: {csv_path}")
        separator = get('defaults.csv_separator', ';')
        df = pd.read_csv(csv_path, sep=separator, quoting=csv.QUOTE_NONNUMERIC,
                         dtype={'SIRET': str, 'SIREN': str, 'NIC': str})
        return csv_path, df

    logger.info("=== ÉTAPE 1: TRAITEMENT DES DONNÉES EXCEL ===")
    start_time = time.time()

    try:
        input_file = params["input_file"]
        csv_path = params["csv_path"]
        sheet_name = params["sheet_name"]

        logger.info(f"Traitement du fichier: {input_file}, feuille: {sheet_name}")
        logger.info(f"Fichier CSV de sortie: {csv_path}")

        df = nettoyer_fichier_excel(input_file, logger, csv_path, sheet_name)

        elapsed = time.time() - start_time
        logger.info(f"Traitement terminé en {elapsed:.2f}s")
        logger.info(f"Données traitées: {len(df)} lignes, {df.columns.size} colonnes")
        logger.info(f"SIRET uniques: {df['SIRET'].nunique()}")

        return csv_path, df

    except Exception as e:
        logger.error(f"Erreur lors du traitement des données: {e}")
        raise


def generate_documents(
    params: Dict[str, Any],
    csv_path: str,
    logger: logging.Logger,
) -> List[str]:
    """
    Génère les attestations DOETH à partir du fichier CSV traité.

    En mode --dry-run, simule la génération sans écrire aucun fichier.
    """
    logger.info("=== ÉTAPE 2: GÉNÉRATION DES ATTESTATIONS ===")
    start_time = time.time()

    if params.get("dry_run"):
        logger.info("[DRY-RUN] Simulation activée — aucun fichier Word/PDF ne sera écrit")
        separator = get('defaults.csv_separator', ';')
        df = pd.read_csv(csv_path, sep=separator, quoting=csv.QUOTE_NONNUMERIC,
                         dtype={'SIRET': str})
        sirets = df['SIRET'].unique()
        logger.info(f"[DRY-RUN] {len(sirets)} attestations seraient générées dans: {params['output_dir']}")
        return [f"[dry-run] {s}" for s in sirets]

    try:
        generated_docs = generer_attestations_doeth(
            csv_path=csv_path,
            output_folder=params["output_dir"],
            session_logger=logger,
            signature_path=params["signature_path"],
            logo_path=params["logo_path"],
            output_format=params.get("output_format", OutputFormat.DOCX),
        )

        elapsed = time.time() - start_time
        logger.info(f"Génération terminée en {elapsed:.2f}s")
        logger.info(f"Attestations générées: {len(generated_docs)}")
        return generated_docs

    except Exception as e:
        logger.error(f"Erreur lors de la génération des attestations: {e}")
        raise


def generate_statistics(df: pd.DataFrame, generated_docs: List[str]) -> RunStats:
    """
    Génère les statistiques structurées sur le traitement effectué.

    Returns:
        RunStats — dataclass typée (remplace le Dict[str, Any] non contraint)
    """
    _log = get_logger("main.generate_statistics")
    _log.info("=== ÉTAPE 3: GÉNÉRATION DES STATISTIQUES ===")

    stats = RunStats()

    try:
        stats.total_rows = len(df)
        stats.unique_sirets = int(df['SIRET'].nunique())
        stats.unique_clients = int(df['NOM_CLIENT'].nunique()) if 'NOM_CLIENT' in df.columns else 0
        stats.total_docs = len(generated_docs)

        if 'REGROUPEMENT' in df.columns:
            stats.regroupements = df['REGROUPEMENT'].value_counts().to_dict()
            _log.info(f"Répartition par regroupement: {stats.regroupements}")

        if 'ETP_ANNUEL' in df.columns:
            stats.total_etp = float(df['ETP_ANNUEL'].sum())
            stats.avg_etp_per_employee = float(df['ETP_ANNUEL'].mean())
            stats.avg_etp_per_siret = float(
                df.groupby('SIRET')['ETP_ANNUEL'].sum().mean())
            _log.info(f"Total ETP: {stats.total_etp:.2f}")

        if 'NB_HEURES' in df.columns:
            stats.total_heures = float(df['NB_HEURES'].sum())
            stats.avg_heures_per_employee = float(df['NB_HEURES'].mean())
            _log.info(f"Total heures: {stats.total_heures:.2f}")

        for doc in generated_docs:
            ext = Path(doc).suffix.lower()
            stats.files_by_format[ext] = stats.files_by_format.get(ext, 0) + 1

        _log.info(f"Types de fichiers générés: {stats.files_by_format}")

    except Exception as e:
        _log.error(f"Erreur lors de la génération des statistiques: {e}")

    return stats


def run_pipeline(
    params: Dict[str, Any],
    logger: logging.Logger,
    progress_callback: Optional[Callable[[float, str], None]] = None,
) -> Tuple[List[str], RunStats]:
    """
    Point d'entrée unique de la pipeline métier.

    Consommable indifféremment par le CLI (main) et le GUI (_worker),
    éliminant la duplication d'orchestration entre les deux interfaces.

    Args:
        params:            Paramètres d'exécution (produits par setup_environment ou le GUI).
        logger:            Logger de session.
        progress_callback: Callback optionnel (pct: float, msg: str) pour la barre de progression GUI.

    Returns:
        Tuple (generated_docs, stats)
    """
    def _progress(pct: float, msg: str = "") -> None:
        if progress_callback:
            progress_callback(pct, msg)

    start_time = time.time()

    _progress(10, "Traitement Excel...")
    csv_path, df_processed = process_data(params, logger)

    _progress(50, "Génération des attestations...")
    generated_docs = generate_documents(params, csv_path, logger)

    _progress(90, "Statistiques...")
    stats = generate_statistics(df_processed, generated_docs)
    stats.duration_seconds = time.time() - start_time

    _progress(100, "Terminé")
    return generated_docs, stats


def main() -> int:
    """Fonction principale d'orchestration du processus."""
    args = parse_arguments()

    try:
        params, logger = setup_environment(args)

        env_name, _ = Config().get_environment()
        logger.info(f"Environnement: {env_name}")

        fmt_map = {
            "docx": OutputFormat.DOCX,
            "pdf": OutputFormat.PDF,
            "both": OutputFormat.BOTH,
        }
        params["output_format"] = fmt_map.get(args.format, OutputFormat.DOCX)

        if params.get("dry_run"):
            logger.info("[DRY-RUN] Mode simulation — aucun fichier Word/PDF ne sera écrit")

        generated_docs, stats = run_pipeline(params, logger)

        logger.info("=== BILAN DU TRAITEMENT ===")
        logger.info(f"Durée totale d'exécution: {stats.duration_seconds:.2f}s")
        logger.info(f"Attestations générées: {stats.total_docs}")
        logger.info(f"SIRET traités: {stats.unique_sirets}")
        logger.info(f"Clients uniques: {stats.unique_clients}")
        if stats.total_etp:
            logger.info(f"Total ETP: {stats.total_etp:.2f}")
        logger.info(f"Dossier de sortie: {params['output_dir']}")
        logger.info("=== TRAITEMENT TERMINÉ AVEC SUCCÈS ===")

        return 0

    except Exception as e:
        logger = get_logger("main")
        logger.error(f"Erreur fatale lors de l'exécution: {e}")
        logger.exception("Détail de l'erreur:")
        logger.error("=== TRAITEMENT TERMINÉ AVEC ERREUR ===")
        return 1


if __name__ == "__main__":
    sys.exit(main())
