"""
Module de traitement des données pour le projet Publipostage DOETH.

Ce module gère le chargement, le nettoyage et la transformation des données
à partir des fichiers Excel source et produit les fichiers CSV intermédiaires.
"""
import csv
import datetime
import logging
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from src.utils.config import get
from src.utils.logger import get_logger

logger = get_logger(__name__)


def load_excel_data(input_file: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    Charge les données depuis un fichier Excel.

    Args:
        input_file: Chemin vers le fichier Excel source
        sheet_name: Nom de la feuille à charger. Si None, utilise la configuration par défaut.

    Returns:
        pd.DataFrame: Les données chargées

    Raises:
        FileNotFoundError: Si le fichier n'existe pas
        ValueError: Si des problèmes surviennent lors du chargement
    """
    if sheet_name is None:
        sheet_name = get('defaults.excel_sheet', 'Feuil1')

    logger.info(f"Chargement du fichier Excel: {input_file}, feuille: {sheet_name}")

    if not Path(input_file).exists():
        error_msg = f"Fichier Excel non trouvé: {input_file}"
        logger.error(error_msg)
        raise FileNotFoundError(error_msg)

    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        logger.info(f"Fichier Excel chargé: {len(df)} lignes, {len(df.columns)} colonnes")
        logger.debug(f"Colonnes disponibles: {', '.join(df.columns)}")
        return df
    except Exception as e:
        error_msg = f"Erreur lors du chargement du fichier Excel: {e}"
        logger.error(error_msg)
        raise ValueError(error_msg) from e


def create_siret_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crée la colonne SIRET en combinant SIREN et NIC.

    Args:
        df: Le DataFrame à modifier

    Returns:
        pd.DataFrame: Le DataFrame avec la colonne SIRET ajoutée

    Raises:
        Exception: Si la construction du SIRET échoue (pipeline arrêté, pas de documents silencieusement vides).
    """
    df_copy = df.copy()

    if 'SIRET' not in df_copy.columns and 'SIREN' in df_copy.columns and 'NIC' in df_copy.columns:
        logger.info("Création de la colonne SIRET à partir de SIREN et NIC")

        df_copy['SIREN'] = df_copy['SIREN'].fillna('').astype(str).str.zfill(9)
        df_copy['NIC'] = df_copy['NIC'].fillna('').astype(str).str.zfill(5)

        invalid_siren_mask = ~df_copy['SIREN'].str.isdigit()
        if invalid_siren_mask.any():
            n = invalid_siren_mask.sum()
            bad_values = df_copy.loc[invalid_siren_mask, 'SIREN'].unique().tolist()
            logger.warning(f"{n} SIREN non-numériques détectés (exclus): {bad_values}")
            df_copy = df_copy[~invalid_siren_mask].copy()

        invalid_nic_mask = ~df_copy['NIC'].str.isdigit()
        if invalid_nic_mask.any():
            n = invalid_nic_mask.sum()
            bad_values = df_copy.loc[invalid_nic_mask, 'NIC'].unique().tolist()
            logger.warning(f"{n} NIC non-numériques détectés (exclus): {bad_values}")
            df_copy = df_copy[~invalid_nic_mask].copy()

        invalid_siren_len = df_copy[df_copy['SIREN'].str.len() != 9]
        if not invalid_siren_len.empty:
            logger.warning(f"{len(invalid_siren_len)} codes SIREN invalides (longueur ≠ 9)")

        invalid_nic_len = df_copy[df_copy['NIC'].str.len() != 5]
        if not invalid_nic_len.empty:
            logger.warning(f"{len(invalid_nic_len)} codes NIC invalides (longueur ≠ 5)")

        df_copy['SIRET'] = df_copy['SIREN'] + df_copy['NIC']
        logger.info(f"{len(df_copy)} codes SIRET générés")

    return df_copy


def format_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Formate les colonnes de dates selon le format configuré.

    Args:
        df: Le DataFrame à modifier

    Returns:
        pd.DataFrame: Le DataFrame avec les dates formatées
    """
    df_copy = df.copy()
    date_format = get('defaults.date_format', '%d/%m/%Y')

    date_columns = [col for col in df_copy.columns if 'DATE' in col.upper()]

    for col in date_columns:
        if col in df_copy.columns:
            logger.info(f"Formatage de la colonne de date: {col} au format {date_format}")
            try:
                df_copy[col] = pd.to_datetime(df_copy[col], errors='coerce')
                null_count = df_copy[col].isna().sum()

                if null_count > 0:
                    logger.warning(f"Colonne {col}: {null_count} valeurs non convertibles en dates")

                df_copy[col] = df_copy[col].dt.strftime(date_format).fillna('')
                logger.info(f"Colonne {col}: dates formatées avec succès")

            except Exception as e:
                logger.error(f"Erreur lors du formatage de la colonne {col}: {e}")

    return df_copy


def clean_and_transform_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Nettoie et transforme les données chargées.

    Args:
        df: Le DataFrame à nettoyer

    Returns:
        pd.DataFrame: Le DataFrame nettoyé
    """
    logger.info("Début du nettoyage et de la transformation des données")

    df_cleaned = df.copy()

    initial_rows = len(df_cleaned)
    df_cleaned.dropna(how='all', inplace=True)
    rows_removed = initial_rows - len(df_cleaned)
    if rows_removed > 0:
        logger.info(f"Suppression de {rows_removed} lignes entièrement vides")

    df_cleaned = create_siret_column(df_cleaned)
    df_cleaned = format_dates(df_cleaned)

    numeric_columns = ['ETP_ANNUEL', 'NB_HEURES']
    for col in numeric_columns:
        if col in df_cleaned.columns:
            try:
                df_cleaned[col] = pd.to_numeric(df_cleaned[col], errors='coerce')
                null_count = df_cleaned[col].isna().sum()
                if null_count > 0:
                    logger.warning(f"Colonne {col}: {null_count} valeurs non numériques → NaN")

                df_cleaned[col] = df_cleaned[col].fillna(0)
                logger.info(f"Colonne {col} nettoyée et convertie en numérique")

            except Exception as e:
                logger.error(f"Erreur lors du nettoyage de la colonne {col}: {e}")

    logger.info(f"Nettoyage terminé: {len(df_cleaned)} lignes conservées")
    return df_cleaned


def aggregate_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrège les données par groupes pertinents.

    Args:
        df: Le DataFrame à agréger

    Returns:
        pd.DataFrame: Le DataFrame agrégé
    """
    logger.info("Agrégation des données")

    group_cols = [
        'CODE_REGROUPEMENT', 'REGROUPEMENT', 'SIREN', 'NIC', 'SIRET',
        'NOM_CLIENT', 'ADRESSE_CLIENT', 'CP_CLIENT', 'VILLE_CLIENT',
        'APE', 'NOM', 'PRENOM', 'DATE_NAISSANCE', 'ANNEE',
        'QUALIFICATION', 'ETP_MAJORE'
    ]

    available_cols = [col for col in group_cols if col in df.columns]
    missing_cols = set(group_cols) - set(available_cols)

    if missing_cols:
        logger.warning(f"Colonnes manquantes pour l'agrégation: {', '.join(missing_cols)}")

    if not available_cols:
        logger.error("Aucune colonne de regroupement disponible. Agrégation impossible.")
        return df

    agg_cols = ['ETP_ANNUEL', 'NB_HEURES']
    available_agg_cols = [col for col in agg_cols if col in df.columns]

    if not available_agg_cols:
        logger.error("Aucune colonne à agréger disponible. Agrégation impossible.")
        return df

    agg_dict = {col: 'sum' for col in available_agg_cols}

    try:
        df_grouped = df.groupby(available_cols, sort=False).agg(agg_dict).reset_index()
        initial_rows = len(df)
        final_rows = len(df_grouped)
        reduction = ((initial_rows - final_rows) / initial_rows * 100) if initial_rows > 0 else 0

        if 'ANNEE' in df_grouped.columns:
            df_grouped['ANNEE'] = df_grouped['ANNEE'].fillna(0).astype(int)

        logger.info(
            f"Données agrégées: {initial_rows} → {final_rows} lignes ({reduction:.1f}% de réduction)")
        return df_grouped

    except Exception as e:
        logger.error(f"Erreur lors de l'agrégation des données: {e}")
        return df


def filter_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filtre les données selon les critères métier.

    Args:
        df: Le DataFrame à filtrer

    Returns:
        pd.DataFrame: Le DataFrame filtré
    """
    logger.info("Filtrage des données selon les critères métier")

    df_filtered = df.copy()
    initial_rows = len(df_filtered)

    if 'CODE_REGROUPEMENT' in df_filtered.columns:
        logger.info("Exclusion des enregistrements avec CODE_REGROUPEMENT = 'DIFFUS'")
        df_filtered = df_filtered[df_filtered['CODE_REGROUPEMENT'] != 'DIFFUS']
        diffus_removed = initial_rows - len(df_filtered)
        logger.info(f"{diffus_removed} enregistrements 'DIFFUS' exclus")

    if 'SIRET' in df_filtered.columns:
        missing_siret = df_filtered['SIRET'].isna().sum()
        if missing_siret > 0:
            logger.warning(f"{missing_siret} enregistrements avec SIRET manquant")

    final_rows = len(df_filtered)
    logger.info(
        f"Filtrage terminé: {initial_rows - final_rows} lignes supprimées, {final_rows} conservées")

    return df_filtered


def add_processing_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ajoute les colonnes nécessaires au traitement par lots.

    Args:
        df: Le DataFrame à enrichir

    Returns:
        pd.DataFrame: Le DataFrame avec les colonnes ajoutées
    """
    logger.info("Ajout des colonnes de traitement pour le regroupement par SIRET")

    df_enhanced = df.copy()

    if 'SIRET' not in df_enhanced.columns:
        logger.error("Impossible d'ajouter les colonnes de traitement: colonne SIRET manquante")
        return df_enhanced

    sort_columns = ['SIRET']
    if 'NOM' in df_enhanced.columns:
        sort_columns.append('NOM')
    if 'PRENOM' in df_enhanced.columns:
        sort_columns.append('PRENOM')

    df_enhanced = df_enhanced.sort_values(by=sort_columns)
    logger.info(f"Données triées par {', '.join(sort_columns)}")

    df_enhanced['NOUVEAU_GROUPE'] = (
        df_enhanced['SIRET'] != df_enhanced['SIRET'].shift(1)).astype(int)
    nouveau_count = df_enhanced['NOUVEAU_GROUPE'].sum()
    logger.info(f"{nouveau_count} groupes SIRET identifiés")

    df_enhanced['FIN_GROUPE'] = (
        df_enhanced['SIRET'] != df_enhanced['SIRET'].shift(-1)).astype(int)
    fin_count = df_enhanced['FIN_GROUPE'].sum()

    if nouveau_count != fin_count:
        logger.warning(
            f"Incohérence: {nouveau_count} débuts de groupe ≠ {fin_count} fins de groupe")
    else:
        logger.info(f"Vérification de cohérence OK: {nouveau_count} groupes SIRET")

    return df_enhanced


def save_processed_data(df: pd.DataFrame, output_file: str) -> str:
    """
    Sauvegarde les données traitées dans un fichier CSV.

    Utilise le pattern write-tmp → rename pour garantir l'atomicité :
    une interruption ne laisse jamais un fichier tronqué à la destination finale.

    Args:
        df: Le DataFrame à sauvegarder
        output_file: Chemin vers le fichier de sortie

    Returns:
        str: Chemin vers le fichier créé
    """
    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    separator = get('defaults.csv_separator', ';')
    tmp_path = output_path.with_suffix('.tmp')

    logger.info(f"Sauvegarde des données traitées: {output_file}")
    try:
        df.to_csv(
            tmp_path,
            index=False,
            sep=separator,
            quoting=csv.QUOTE_NONNUMERIC,
            encoding='utf-8',
        )
        tmp_path.rename(output_path)

        file_size = output_path.stat().st_size / 1024
        logger.info(f"Fichier CSV créé: {len(df)} lignes, {file_size:.1f} Ko")
        return str(output_path)

    except Exception:
        if tmp_path.exists():
            tmp_path.unlink()
        raise


def nettoyer_fichier_excel(
    input_file: str,
    logger: logging.Logger,
    output_file: Optional[str] = None,
    sheet_name: Optional[str] = None,
) -> pd.DataFrame:
    """
    Fonction principale qui orchestre le traitement complet des données.

    Args:
        input_file: Chemin vers le fichier Excel source
        logger: Logger à utiliser pour les messages
        output_file: Chemin vers le fichier CSV de sortie. Si None, génère un nom automatique.
        sheet_name: Nom de la feuille à traiter. Si None, utilise celle de la configuration.

    Returns:
        pd.DataFrame: Le DataFrame final traité
    """
    logger.info(f"=== DÉBUT DU TRAITEMENT DU FICHIER: {input_file} ===")

    if output_file is None:
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        processed_dir = Path(get('paths.processed_dir', './data/processed'))
        processed_dir.mkdir(parents=True, exist_ok=True)
        output_file = str(processed_dir / f"processed_{timestamp}.csv")
        logger.info(f"Nom de fichier de sortie généré: {output_file}")

    try:
        logger.info("ÉTAPE 1: Chargement des données Excel")
        df = load_excel_data(input_file, sheet_name)

        logger.info("ÉTAPE 2: Nettoyage et transformation des données")
        df_cleaned = clean_and_transform_data(df)

        logger.info("ÉTAPE 3: Agrégation des données")
        df_grouped = aggregate_data(df_cleaned)

        logger.info("ÉTAPE 4: Filtrage des données")
        df_filtered = filter_data(df_grouped)

        logger.info("ÉTAPE 5: Ajout des colonnes de traitement")
        df_final = add_processing_columns(df_filtered)

        logger.info("ÉTAPE 6: Sauvegarde des données traitées")
        save_processed_data(df_final, output_file)

        logger.info(f"=== TRAITEMENT TERMINÉ AVEC SUCCÈS. FICHIER CRÉÉ: {output_file} ===")
        return df_final

    except Exception as e:
        logger.error(f"!!! ERREUR LORS DU TRAITEMENT DES DONNÉES: {e}")
        raise
