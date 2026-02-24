#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Point d'entrée principal du projet Publipostage DOETH.

Ce script orchestre l'ensemble du processus de traitement des données
et de génération des attestations DOETH.
"""

import os
import sys
import time
import argparse
import datetime
import pandas as pd
import logging
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple

from src.utils.config import Config, get
from src.utils.logger import setup_logger, get_logger
from src.data_processor import nettoyer_fichier_excel
from src.document_generator import generer_attestations_doeth


def parse_arguments():
    """
    Parse les arguments de ligne de commande.

    Returns:
        argparse.Namespace: Les arguments parsés
    """
    parser = argparse.ArgumentParser(
        description="Publipostage DOETH - Génération d'attestations à partir de fichiers Excel"
    )

    parser.add_argument(
        "--config",
        type=str,
        help="Chemin vers un fichier de configuration personnalisé"
    )

    parser.add_argument(
        "--input",
        type=str,
        help="Chemin vers le fichier Excel d'entrée (override config)"
    )

    parser.add_argument(
        "--sheet",
        type=str,
        help="Nom de la feuille Excel à traiter (override config)"
    )

    parser.add_argument(
        "--output-dir",
        type=str,
        help="Dossier de sortie pour les attestations (override config)"
    )

    parser.add_argument(
        "--skip-processing",
        action="store_true",
        help="Ignorer l'étape de traitement de l'Excel et utiliser le CSV déjà généré"
    )

    parser.add_argument(
        "--csv-path",
        type=str,
        help="Chemin vers un fichier CSV déjà traité (utilisé avec --skip-processing)"
    )

    parser.add_argument(
        "--debug",
        action="store_true",
        help="Active le mode debug avec logs détaillés"
    )

    parser.add_argument(
        "--format",
        choices=["docx", "pdf", "both"],
        default="docx",
        help="Format de sortie des attestations : docx (défaut), pdf, both"
    )

    return parser.parse_args()


def setup_environment(args) -> Dict[str, Any]:
    """
    Configure l'environnement d'exécution et initialise les ressources.

    Args:
        args: Arguments de ligne de commande

    Returns:
        Dict[str, Any]: Dictionnaire contenant les paramètres d'exécution
    """
    # Déterminer le niveau de log
    console_level = logging.DEBUG if args.debug else logging.INFO

    # Configurer le logger
    logs_dir = get('paths.logs_dir')
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    logger = setup_logger(
        logs_dir=logs_dir,
        name=f"publipostage_{timestamp}",
        console_level=console_level,
        file_level=logging.DEBUG,
        enable_colors=True
    )

    logger.info("=== DÉMARRAGE DU PUBLIPOSTAGE DOETH ===")
    logger.info(
        f"Date et heure: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    # Collecter les paramètres
    params = {
        "timestamp": timestamp,
        "log_level": console_level,
        "logs_dir": logs_dir,
    }

    # Déterminer les chemins de fichiers
    if args.input:
        params["input_file"] = args.input
    else:
        input_dir = get('paths.input_dir')
        input_filename = get('defaults.input_filename')
        params["input_file"] = os.path.join(input_dir, input_filename)

    # Déterminer la feuille Excel
    if args.sheet:
        params["sheet_name"] = args.sheet
    else:
        params["sheet_name"] = get('defaults.excel_sheet')

    # Déterminer le dossier de sortie
    if args.output_dir:
        params["output_dir"] = args.output_dir
    else:
        params["output_dir"] = get('paths.output_dir')

    # Déterminer si on ignore l'étape de traitement
    params["skip_processing"] = args.skip_processing

    # Déterminer le chemin du CSV traité
    if args.csv_path:
        params["csv_path"] = args.csv_path
    else:
        processed_dir = get('paths.processed_dir')
        params["csv_path"] = os.path.join(
            processed_dir, f"processed_{timestamp}.csv")

    # Vérification des chemins et création des dossiers nécessaires
    dirs_to_check = [
        logs_dir,
        get('paths.input_dir'),
        get('paths.processed_dir'),
        get('paths.output_dir'),
    ]

    for dir_path in dirs_to_check:
        if not os.path.exists(dir_path):
            logger.info(f"Création du dossier: {dir_path}")
            os.makedirs(dir_path, exist_ok=True)

    # Vérification des ressources
    params["logo_path"] = get('resources.logo_path')
    params["signature_path"] = get('resources.signature_path')

    for resource_name, resource_path in [
        ("Logo", params["logo_path"]),
        ("Signature", params["signature_path"])
    ]:
        if not resource_path or not os.path.exists(resource_path):
            logger.warning(f"{resource_name} non trouvé: {resource_path}")
        else:
            logger.info(f"{resource_name} trouvé: {resource_path}")

    # Vérification du fichier d'entrée
    if not params["skip_processing"] and not os.path.exists(params["input_file"]):
        logger.error(f"Fichier d'entrée non trouvé: {params['input_file']}")
        raise FileNotFoundError(
            f"Fichier d'entrée non trouvé: {params['input_file']}")

    logger.info("Environnement d'exécution configuré avec succès")
    return params


def process_data(params: Dict[str, Any], logger: logging.Logger) -> str:
    """
    Traite les données Excel pour préparer le fichier CSV intermédiaire.

    Args:
        params: Paramètres d'exécution
        logger:

    Returns:
        str: Chemin vers le fichier CSV traité
    """

    if params["skip_processing"]:
        csv_path = params["csv_path"]
        if not os.path.exists(csv_path):
            logger.error(f"Fichier CSV spécifié non trouvé: {csv_path}")
            raise FileNotFoundError(
                f"Fichier CSV spécifié non trouvé: {csv_path}")

        logger.info(
            f"Étape de traitement ignorée, utilisation du CSV existant: {csv_path}")
        return csv_path

    logger.info("=== ÉTAPE 1: TRAITEMENT DES DONNÉES EXCEL ===")

    start_time = time.time()

    try:
        input_file = params["input_file"]
        csv_path = params["csv_path"]
        sheet_name = params["sheet_name"]

        logger.info(
            f"Traitement du fichier: {input_file}, feuille: {sheet_name}")
        logger.info(f"Fichier CSV de sortie: {csv_path}")

        # Appel à la fonction de traitement des données
        df_processed = nettoyer_fichier_excel(
            input_file, logger, csv_path, sheet_name)

        elapsed_time = time.time() - start_time
        logger.info(f"Traitement terminé en {elapsed_time:.2f} secondes")
        logger.info(
            f"Données traitées: {len(df_processed)} lignes, {df_processed.columns.size} colonnes")
        logger.info(
            f"Nombre de SIRET uniques: {df_processed['SIRET'].nunique()}")

        return csv_path

    except Exception as e:
        logger.error(f"Erreur lors du traitement des données: {str(e)}")
        raise


def generate_documents(params: Dict[str, Any], csv_path: str, logger: logging.Logger) -> List[str]:
    """
    Génère les attestations DOETH à partir du fichier CSV traité.

    Args:
        params: Paramètres d'exécution
        csv_path: Chemin vers le fichier CSV traité
        logger:

    Returns:
        List[str]: Liste des chemins des attestations générées
    """
    logger.info("=== ÉTAPE 2: GÉNÉRATION DES ATTESTATIONS ===")

    start_time = time.time()

    try:
        output_dir = params["output_dir"]
        logo_path = params["logo_path"]
        signature_path = params["signature_path"]

        logger.info(f"Génération des attestations à partir de: {csv_path}")
        logger.info(f"Dossier de sortie: {output_dir}")

        # Appel à la fonction de génération des attestations
        generated_docs = generer_attestations_doeth(
            csv_path=csv_path,
            output_folder=output_dir,
            logger=logger,
            signature_path=signature_path,
            logo_path=logo_path,
            output_format=params.get("output_format"),
        )

        elapsed_time = time.time() - start_time
        logger.info(f"Génération terminée en {elapsed_time:.2f} secondes")
        logger.info(f"Nombre d'attestations générées: {len(generated_docs)}")

        return generated_docs

    except Exception as e:
        logger.error(
            f"Erreur lors de la génération des attestations: {str(e)}")
        raise


def generate_statistics(df: pd.DataFrame, generated_docs: List[str]) -> Dict[str, Any]:
    """
    Génère des statistiques sur le traitement effectué.

    Args:
        df: DataFrame des données traitées
        generated_docs: Liste des documents générés

    Returns:
        Dict[str, Any]: Dictionnaire contenant les statistiques
    """
    logger = get_logger("main.generate_statistics")
    logger.info("=== ÉTAPE 3: GÉNÉRATION DES STATISTIQUES ===")

    stats = {}

    try:
        # Statistiques de base
        stats["total_rows"] = len(df)
        stats["unique_sirets"] = df['SIRET'].nunique()
        stats["unique_clients"] = df['NOM_CLIENT'].nunique(
        ) if 'NOM_CLIENT' in df.columns else 0
        stats["total_docs"] = len(generated_docs)

        # Statistiques par regroupement si la colonne existe
        if 'REGROUPEMENT' in df.columns:
            regroupements = df['REGROUPEMENT'].value_counts().to_dict()
            stats["regroupements"] = regroupements
            logger.info(f"Répartition par regroupement: {regroupements}")

        # Statistiques sur les ETP
        if 'ETP_ANNUEL' in df.columns:
            stats["total_etp"] = df['ETP_ANNUEL'].sum()
            stats["avg_etp_per_employee"] = df['ETP_ANNUEL'].mean()
            stats["avg_etp_per_siret"] = df.groupby(
                'SIRET')['ETP_ANNUEL'].sum().mean()
            logger.info(f"Total ETP: {stats['total_etp']:.2f}")
            logger.info(
                f"Moyenne ETP par employé: {stats['avg_etp_per_employee']:.2f}")
            logger.info(
                f"Moyenne ETP par SIRET: {stats['avg_etp_per_siret']:.2f}")

        # Statistiques sur les heures
        if 'NB_HEURES' in df.columns:
            stats["total_heures"] = df['NB_HEURES'].sum()
            stats["avg_heures_per_employee"] = df['NB_HEURES'].mean()
            logger.info(f"Total heures: {stats['total_heures']:.2f}")
            logger.info(
                f"Moyenne heures par employé: {stats['avg_heures_per_employee']:.2f}")

        # Autres statistiques
        stats["file_count_by_extension"] = {}
        for doc in generated_docs:
            ext = os.path.splitext(doc)[1].lower()
            if ext in stats["file_count_by_extension"]:
                stats["file_count_by_extension"][ext] += 1
            else:
                stats["file_count_by_extension"][ext] = 1

        logger.info(
            f"Types de fichiers générés: {stats['file_count_by_extension']}")

        return stats

    except Exception as e:
        logger.error(
            f"Erreur lors de la génération des statistiques: {str(e)}")
        # En cas d'erreur, on renvoie tout de même les statistiques partielles
        return stats


def main():
    """
    Fonction principale d'orchestration du processus.
    """
    # Parsing des arguments
    args = parse_arguments()

    try:
        # Configuration de l'environnement
        params = setup_environment(args)

        # Déterminer l'environnement d'exécution
        env_name, _ = Config().get_environment()

        # Créer un logger spécifique pour la session principale
        logs_dir = get('paths.logs_dir')
        logger = setup_logger(
            logs_dir=logs_dir,
            name=f"publipostage_doeth_{env_name.lower()}_{datetime.datetime.now().strftime('%Y%m%d')}",
            console_level=params["log_level"],
            file_level=logging.DEBUG,
            enable_colors=True
        )

        start_time = time.time()

        # Étape 1: Traitement des données
        csv_path = process_data(params, logger)

        # Charger le DataFrame pour les statistiques finales
        separator = get('defaults.csv_separator', ';')
        df_processed = pd.read_csv(csv_path, sep=separator, quoting=1)

        # Étape 2: Génération des attestations
        from src.document_generator import OutputFormat
        fmt_map = {"docx": OutputFormat.DOCX,
                   "pdf": OutputFormat.PDF, "both": OutputFormat.BOTH}
        params["output_format"] = fmt_map.get(args.format, OutputFormat.DOCX)
        generated_docs = generate_documents(params, csv_path, logger)

        # Étape 3: Génération des statistiques
        stats = generate_statistics(df_processed, generated_docs)

        # Bilan final
        total_time = time.time() - start_time
        logger.info("=== BILAN DU TRAITEMENT ===")
        logger.info(f"Durée totale d'exécution: {total_time:.2f} secondes")
        logger.info(
            f"Nombre total d'attestations générées: {len(generated_docs)}")
        logger.info(f"Nombre de SIRET traités: {stats['unique_sirets']}")
        logger.info(f"Nombre de clients uniques: {stats['unique_clients']}")

        if 'total_etp' in stats:
            logger.info(
                f"Total d'unités bénéficiaires (ETP): {stats['total_etp']:.2f}")

        logger.info(f"Dossier de sortie: {params['output_dir']}")
        logger.info("=== TRAITEMENT TERMINÉ AVEC SUCCÈS ===")

        return 0

    except Exception as e:
        logger = get_logger("main")
        logger.error(f"Erreur fatale lors de l'exécution: {str(e)}")
        logger.exception("Détail de l'erreur:")
        logger.error("=== TRAITEMENT TERMINÉ AVEC ERREUR ===")
        return 1


if __name__ == "__main__":
    sys.exit(main())
