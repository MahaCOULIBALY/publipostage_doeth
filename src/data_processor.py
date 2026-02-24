"""
Module de traitement des données pour le projet Publipostage DOETH.

Ce module gère le chargement, le nettoyage et la transformation des données
à partir des fichiers Excel source et produit les fichiers CSV intermédiaires.
"""
import logging
import os
import pandas as pd
import numpy as np
import csv
import datetime
from pathlib import Path
from typing import Optional, Union, Dict, Any, Tuple

from src.utils.config import get
from src.utils.logger import get_logger

# Initialisation du logger
logger = get_logger(__name__)


def load_excel_data(input_file: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    Charge les données depuis un fichier Excel.

    Args:
        input_file (str): Chemin vers le fichier Excel source
        sheet_name (str, optional): Nom de la feuille à charger. Si None, utilise la configuration par défaut.

    Returns:
        pd.DataFrame: Les données chargées

    Raises:
        FileNotFoundError: Si le fichier n'existe pas
        ValueError: Si des problèmes surviennent lors du chargement
    """
    if sheet_name is None:
        sheet_name = get('defaults.excel_sheet', 'Feuil1')

    logger.info(
        f"Chargement du fichier Excel: {input_file}, feuille: {sheet_name}")

    if not os.path.exists(input_file):
        error_msg = f"Fichier Excel non trouvé: {input_file}"
        logger.error(error_msg)
        raise FileNotFoundError(error_msg)

    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        row_count = len(df)
        col_count = len(df.columns)
        logger.info(
            f"Fichier Excel chargé avec succès: {row_count} lignes, {col_count} colonnes")

        # Afficher un aperçu des colonnes pour le débogage
        logger.debug(f"Colonnes disponibles: {', '.join(df.columns)}")

        return df
    except Exception as e:
        error_msg = f"Erreur lors du chargement du fichier Excel: {str(e)}"
        logger.error(error_msg)
        raise ValueError(error_msg)


def create_siret_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crée la colonne SIRET en combinant SIREN et NIC.

    Args:
        df (pd.DataFrame): Le DataFrame à modifier

    Returns:
        pd.DataFrame: Le DataFrame avec la colonne SIRET ajoutée
    """
    df_copy = df.copy()

    if 'SIRET' not in df_copy.columns and 'SIREN' in df_copy.columns and 'NIC' in df_copy.columns:
        logger.info("Création de la colonne SIRET à partir de SIREN et NIC")

        try:
            # Convertir en chaînes de caractères, gérer les valeurs manquantes et le formatage des codes SIREN sur 9 digit et NIC sur 5 digit
            df_copy['SIREN'] = df_copy['SIREN'].fillna(
                '').astype(str).str.zfill(9)
            df_copy['NIC'] = df_copy['NIC'].fillna('').astype(str).str.zfill(5)

            # Vérifier que SIREN et NIC sont purement numériques (longueur ET contenu)
            invalid_siren_mask = ~df_copy['SIREN'].str.isdigit()
            if invalid_siren_mask.any():
                n = invalid_siren_mask.sum()
                bad_values = df_copy.loc[invalid_siren_mask,
                                         'SIREN'].unique().tolist()
                logger.warning(
                    f"{n} SIREN non-numériques détectés (exclus): {bad_values}")
                df_copy = df_copy[~invalid_siren_mask].copy()

            invalid_nic_mask = ~df_copy['NIC'].str.isdigit()
            if invalid_nic_mask.any():
                n = invalid_nic_mask.sum()
                bad_values = df_copy.loc[invalid_nic_mask,
                                         'NIC'].unique().tolist()
                logger.warning(
                    f"{n} NIC non-numériques détectés (exclus): {bad_values}")
                df_copy = df_copy[~invalid_nic_mask].copy()

            # Vérifier la longueur après nettoyage
            invalid_siren_len = df_copy[df_copy['SIREN'].str.len() != 9]
            if not invalid_siren_len.empty:
                logger.warning(
                    f"Détection de {len(invalid_siren_len)} codes SIREN invalides (longueur ≠ 9)")

            invalid_nic_len = df_copy[df_copy['NIC'].str.len() != 5]
            if not invalid_nic_len.empty:
                logger.warning(
                    f"Détection de {len(invalid_nic_len)} codes NIC invalides (longueur ≠ 5)")

            # Créer la colonne SIRET (14 chiffres = 9 SIREN + 5 NIC)
            df_copy['SIRET'] = df_copy['SIREN'] + df_copy['NIC']
            logger.info(f"{len(df_copy)} codes SIRET générés")

        except Exception as e:
            logger.error(
                f"Erreur lors de la création de la colonne SIRET: {str(e)}")
            # En cas d'erreur, nous continuons sans la colonne SIRET

    return df_copy


def format_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Formate les colonnes de dates selon le format configuré.

    Args:
        df (pd.DataFrame): Le DataFrame à modifier

    Returns:
        pd.DataFrame: Le DataFrame avec les dates formatées
    """
    df_copy = df.copy()
    date_format = get('defaults.date_format', '%d/%m/%Y')

    # Liste des colonnes contenant potentiellement des dates
    date_columns = [col for col in df_copy.columns if 'DATE' in col.upper()]

    for col in date_columns:
        if col in df_copy.columns:
            logger.info(
                f"Formatage de la colonne de date: {col} au format {date_format}")
            try:
                # Convertir en datetime puis au format souhaité
                df_copy[col] = pd.to_datetime(df_copy[col], errors='coerce')
                non_null_count = df_copy[col].count()
                null_count = df_copy[col].isna().sum()

                if null_count > 0:
                    logger.warning(
                        f"Colonne {col}: {null_count} valeurs non convertibles en dates")

                # Appliquer le format de date configuré
                df_copy[col] = df_copy[col].dt.strftime(date_format).fillna('')
                logger.info(
                    f"Colonne {col}: {non_null_count} dates formatées avec succès")

            except Exception as e:
                logger.error(
                    f"Erreur lors du formatage de la colonne {col}: {str(e)}")

    return df_copy


def clean_and_transform_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Nettoie et transforme les données chargées.

    Opérations réalisées:
    - Création de la colonne SIRET
    - Formatage des dates
    - Nettoyage général des données

    Args:
        df (pd.DataFrame): Le DataFrame à nettoyer

    Returns:
        pd.DataFrame: Le DataFrame nettoyé
    """
    logger.info("Début du nettoyage et de la transformation des données")

    # Créer une copie pour éviter de modifier l'original
    df_cleaned = df.copy()

    # Supprimer les lignes entièrement vides
    initial_rows = len(df_cleaned)
    df_cleaned.dropna(how='all', inplace=True)
    rows_removed = initial_rows - len(df_cleaned)
    if rows_removed > 0:
        logger.info(f"Suppression de {rows_removed} lignes entièrement vides")

    # Création de la colonne SIRET
    df_cleaned = create_siret_column(df_cleaned)

    # Formatage des dates
    df_cleaned = format_dates(df_cleaned)

    # Nettoyer les valeurs aberrantes dans les colonnes numériques
    numeric_columns = ['ETP_ANNUEL', 'NB_HEURES']
    for col in numeric_columns:
        if col in df_cleaned.columns:
            try:
                # Convertir en numérique, forçant les valeurs non numériques à NaN
                df_cleaned[col] = pd.to_numeric(
                    df_cleaned[col], errors='coerce')
                null_count = df_cleaned[col].isna().sum()
                if null_count > 0:
                    logger.warning(
                        f"Colonne {col}: {null_count} valeurs non numériques remplacées par NaN")

                # Remplacer les NaN par 0 (correction pour éviter le warning avec inplace=True)
                df_cleaned[col] = df_cleaned[col].fillna(0)
                logger.info(
                    f"Colonne {col} nettoyée et convertie en numérique")

            except Exception as e:
                logger.error(
                    f"Erreur lors du nettoyage de la colonne {col}: {str(e)}")

    logger.info(f"Nettoyage terminé: {len(df_cleaned)} lignes conservées")
    return df_cleaned


def aggregate_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrège les données par groupes pertinents.

    Args:
        df (pd.DataFrame): Le DataFrame à agréger

    Returns:
        pd.DataFrame: Le DataFrame agrégé
    """
    logger.info("Agrégation des données")

    # Colonnes clés pour le regroupement
    group_cols = [
        'CODE_REGROUPEMENT', 'REGROUPEMENT', 'SIREN', 'NIC', 'SIRET',
        'NOM_CLIENT', 'ADRESSE_CLIENT', 'CP_CLIENT', 'VILLE_CLIENT',
        'APE', 'NOM', 'PRENOM', 'DATE_NAISSANCE', 'ANNEE',
        'QUALIFICATION', 'ETP_MAJORE'
    ]

    # Vérifier quelles colonnes de regroupement sont disponibles
    available_cols = [col for col in group_cols if col in df.columns]
    missing_cols = set(group_cols) - set(available_cols)

    if missing_cols:
        logger.warning(
            f"Colonnes manquantes pour l'agrégation: {', '.join(missing_cols)}")

    if not available_cols:
        logger.error(
            "Aucune colonne de regroupement disponible. Agrégation impossible.")
        return df

    # Colonnes à agréger
    agg_cols = ['ETP_ANNUEL', 'NB_HEURES']
    available_agg_cols = [col for col in agg_cols if col in df.columns]

    if not available_agg_cols:
        logger.error(
            "Aucune colonne à agréger disponible. Agrégation impossible.")
        return df

    # Créer le dictionnaire d'agrégation (somme pour toutes les colonnes numériques)
    agg_dict = {col: 'sum' for col in available_agg_cols}

    # Effectuer l'agrégation
    try:
        df_grouped = df.groupby(available_cols, sort=False).agg(
            agg_dict).reset_index()
        initial_rows = len(df)
        final_rows = len(df_grouped)
        reduction = ((initial_rows - final_rows) /
                     initial_rows * 100) if initial_rows > 0 else 0

        # Restaurer le type entier d'ANNEE (upcaste en float64 pendant le groupby si NaN présents)
        if 'ANNEE' in df_grouped.columns:
            df_grouped['ANNEE'] = df_grouped['ANNEE'].fillna(0).astype(int)

        logger.info(
            f"Données agrégées: {initial_rows} → {final_rows} lignes ({reduction:.1f}% de réduction)")
        return df_grouped

    except Exception as e:
        logger.error(f"Erreur lors de l'agrégation des données: {str(e)}")
        return df


def filter_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filtre les données selon les critères métier.

    Args:
        df (pd.DataFrame): Le DataFrame à filtrer

    Returns:
        pd.DataFrame: Le DataFrame filtré
    """
    logger.info("Filtrage des données selon les critères métier")

    # Créer une copie pour éviter de modifier l'original
    df_filtered = df.copy()
    initial_rows = len(df_filtered)

    # Filtrer les DIFFUS si la colonne existe
    if 'CODE_REGROUPEMENT' in df_filtered.columns:
        logger.info(
            "Exclusion des enregistrements avec CODE_REGROUPEMENT = 'DIFFUS'")
        df_filtered = df_filtered[df_filtered['CODE_REGROUPEMENT'] != 'DIFFUS']
        diffus_removed = initial_rows - len(df_filtered)
        logger.info(f"{diffus_removed} enregistrements 'DIFFUS' exclus")

    # Vérifier s'il y a des SIRET manquants ou invalides
    if 'SIRET' in df_filtered.columns:
        missing_siret = df_filtered['SIRET'].isna().sum()
        if missing_siret > 0:
            logger.warning(
                f"{missing_siret} enregistrements avec SIRET manquant")
            # Option: filtrer les lignes sans SIRET
            # df_filtered = df_filtered[df_filtered['SIRET'].notna()]

    # Autres filtres métier pourraient être ajoutés ici

    final_rows = len(df_filtered)
    removed_rows = initial_rows - final_rows
    logger.info(
        f"Filtrage terminé: {removed_rows} lignes supprimées, {final_rows} lignes conservées")

    return df_filtered


def add_processing_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ajoute les colonnes nécessaires au traitement par lots.

    Args:
        df (pd.DataFrame): Le DataFrame à enrichir

    Returns:
        pd.DataFrame: Le DataFrame avec les colonnes ajoutées
    """
    logger.info(
        "Ajout des colonnes de traitement pour le regroupement par SIRET")

    # Créer une copie pour éviter de modifier l'original
    df_enhanced = df.copy()

    # Vérifier que la colonne SIRET existe
    if 'SIRET' not in df_enhanced.columns:
        logger.error(
            "Impossible d'ajouter les colonnes de traitement: colonne SIRET manquante")
        return df_enhanced

    # Tri par SIRET pour assurer la cohérence du traitement par lots
    sort_columns = ['SIRET']

    # Ajouter NOM et PRENOM au tri si disponibles
    if 'NOM' in df_enhanced.columns:
        sort_columns.append('NOM')
    if 'PRENOM' in df_enhanced.columns:
        sort_columns.append('PRENOM')

    df_enhanced = df_enhanced.sort_values(by=sort_columns)
    logger.info(f"Données triées par {', '.join(sort_columns)}")

    # Ajout de la colonne NOUVEAU_GROUPE (1 pour le premier employé de chaque SIRET, 0 pour les autres)
    logger.info("Ajout de la colonne NOUVEAU_GROUPE")
    df_enhanced['NOUVEAU_GROUPE'] = (
        df_enhanced['SIRET'] != df_enhanced['SIRET'].shift(1)).astype(int)
    nouveau_count = df_enhanced['NOUVEAU_GROUPE'].sum()
    logger.info(f"{nouveau_count} groupes SIRET identifiés")

    # Ajout de la colonne FIN_GROUPE (1 pour le dernier employé de chaque SIRET, 0 pour les autres)
    logger.info("Ajout de la colonne FIN_GROUPE")
    df_enhanced['FIN_GROUPE'] = (
        df_enhanced['SIRET'] != df_enhanced['SIRET'].shift(-1)).astype(int)
    fin_count = df_enhanced['FIN_GROUPE'].sum()

    # Vérification de cohérence
    if nouveau_count != fin_count:
        logger.warning(
            f"Incohérence détectée: {nouveau_count} débuts de groupe ≠ {fin_count} fins de groupe")
    else:
        logger.info(
            f"Vérification de cohérence OK: {nouveau_count} groupes SIRET bien identifiés")

    return df_enhanced


def save_processed_data(df: pd.DataFrame, output_file: str) -> str:
    """
    Sauvegarde les données traitées dans un fichier CSV.

    Args:
        df (pd.DataFrame): Le DataFrame à sauvegarder
        output_file (str): Chemin vers le fichier de sortie

    Returns:
        str: Chemin vers le fichier créé
    """
    # Créer le répertoire de sortie si nécessaire
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
        logger.info(f"Création du répertoire de sortie: {output_dir}")

    # Déterminer le séparateur à utiliser
    separator = get('defaults.csv_separator', ';')

    logger.info(f"Sauvegarde des données traitées: {output_file}")
    try:
        # Sauvegarde avec paramètres optimaux pour le publipostage
        df.to_csv(
            output_file,
            index=False,
            sep=separator,
            quoting=csv.QUOTE_NONNUMERIC,
            encoding='utf-8'
        )

        # Vérification que le fichier a bien été créé
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file) / 1024  # Taille en Ko
            logger.info(
                f"Fichier CSV créé avec succès: {len(df)} lignes, {file_size:.1f} Ko")
            return output_file
        else:
            logger.error(
                f"Échec de vérification: le fichier {output_file} n'existe pas après sauvegarde")
            raise FileNotFoundError(
                f"Le fichier {output_file} n'a pas été créé")

    except Exception as e:
        logger.error(f"Erreur lors de la sauvegarde du fichier CSV: {str(e)}")
        raise


def nettoyer_fichier_excel(input_file: str, logger: logging.Logger, output_file: str = None, sheet_name: str = None) -> pd.DataFrame:
    """
    Fonction principale qui orchestre le traitement complet des données.

    Args:
        input_file (str): Chemin vers le fichier Excel source
        output_file (str, optional): Chemin vers le fichier CSV de sortie. Si None, génère un nom automatique.
        sheet_name (str, optional): Nom de la feuille à traiter. Si None, utilise celle de la configuration.
        logger (logging.Logger): Logger à utiliser pour les messages

    Returns:
        pd.DataFrame: Le DataFrame final traité
    """
    logger.info(f"=== DÉBUT DU TRAITEMENT DU FICHIER: {input_file} ===")

    # Générer un nom de fichier de sortie si non spécifié
    if output_file is None:
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        processed_dir = get('paths.processed_dir', './data/processed')
        os.makedirs(processed_dir, exist_ok=True)
        output_file = os.path.join(processed_dir, f"processed_{timestamp}.csv")
        logger.info(f"Nom de fichier de sortie généré: {output_file}")

    try:
        # ÉTAPE 1: Chargement des données
        logger.info("ÉTAPE 1: Chargement des données Excel")
        df = load_excel_data(input_file, sheet_name)

        # ÉTAPE 2: Nettoyage et transformation
        logger.info("ÉTAPE 2: Nettoyage et transformation des données")
        df_cleaned = clean_and_transform_data(df)

        # ÉTAPE 3: Agrégation des données
        logger.info("ÉTAPE 3: Agrégation des données")
        df_grouped = aggregate_data(df_cleaned)

        # ÉTAPE 4: Filtrage des données
        logger.info("ÉTAPE 4: Filtrage des données")
        df_filtered = filter_data(df_grouped)

        # ÉTAPE 5: Ajout des colonnes de traitement
        logger.info("ÉTAPE 5: Ajout des colonnes de traitement")
        df_final = add_processing_columns(df_filtered)

        # ÉTAPE 6: Sauvegarde des données traitées
        logger.info("ÉTAPE 6: Sauvegarde des données traitées")
        save_processed_data(df_final, output_file)

        logger.info(
            f"=== TRAITEMENT TERMINÉ AVEC SUCCÈS. FICHIER CRÉÉ: {output_file} ===")
        return df_final

    except Exception as e:
        logger.error(f"!!! ERREUR LORS DU TRAITEMENT DES DONNÉES: {str(e)}")
        # Remonter l'exception pour qu'elle soit gérée au niveau supérieur
        raise
