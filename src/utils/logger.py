# src/utils/logger.py
import logging
import sys
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path
import inspect
from typing import Optional, Union, Dict

import colorama


class ColorFormatter(logging.Formatter):
    """Formatter personnalisé pour l'affichage coloré des logs dans la console"""

    if sys.platform.lower() == 'win32':
        os.system('color')  # Activation du support ANSI sur Windows

    COLORS = {
        'DEBUG': '\x1b[38;5;245m',  # Gris
        'INFO': '\x1b[38;5;15m',  # Blanc
        'WARNING': '\x1b[38;5;214m',  # Orange
        'ERROR': '\x1b[38;5;196m',  # Rouge vif
        'CRITICAL': '\x1b[97;41m'  # Texte blanc sur fond rouge
    }
    RESET = '\x1b[0m'

    def format(self, record):
        original_msg = record.msg
        levelname = record.levelname
        color = self.COLORS.get(levelname, self.RESET)

        # Application de la couleur au message
        record.msg = f"{color}{original_msg}{self.RESET}"
        formatted_message = super().format(record)

        # Restauration du message original
        record.msg = original_msg
        return f"{color}{formatted_message}{self.RESET}"


class FunctionNameFilter(logging.Filter):
    """Filtre pour inclure le nom de la fonction appelante dans les logs"""

    def filter(self, record):
        # Parcours de la stack d'appels pour trouver la fonction appelante
        for frame in inspect.stack()[1:]:
            if 'logging' not in frame.filename:
                record.funcName = frame.function
                break
        else:
            record.funcName = 'Unknown'
        return True


class SafeRotatingFileHandler(RotatingFileHandler):
    """Gestion robuste des fichiers de log avec création automatique des répertoires"""

    def __init__(self, filename: Union[str, Path], **kwargs):
        # Conversion en Path et création des répertoires
        path = Path(filename).absolute()
        path.parent.mkdir(parents=True, exist_ok=True)

        # Conversion du Path en string pour le parent
        super().__init__(str(path), **kwargs)


def setup_logger(logs_dir: Union[str, Path],
    name: str = "app_logger",
    *,
    console_level: int = logging.INFO,
    file_level: int = logging.DEBUG,
    max_bytes: int = 10 * 1024 * 1024,  # 10 MB
    backup_count: int = 10,
    log_format: Optional[str] = None,
    date_format: str = "%Y-%m-%d %H:%M:%S",
    enable_colors: bool = True) -> logging.Logger:
    """
    Configure et retourne un logger avec rotation des fichiers et formatage coloré pour la console.

    Args:
        logs_dir: Répertoire de stockage des logs
        name: Nom unique du logger
        console_level: Niveau de log pour la console
        file_level: Niveau de log pour les fichiers
        max_bytes: Taille max des fichiers avant rotation
        backup_count: Nombre de fichiers de backup à conserver
        log_format: Format personnalisé des logs
        date_format: Format de date/heure
        enable_colors: Active les couleurs dans la console

    Returns:
        logging.Logger: Logger configuré avec handlers console et fichier
    """
    # Initialisation du logger
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger  # Évite la réinitialisation multiple

    logger.setLevel(min(console_level, file_level))
    logger.propagate = False

    # Ajout du filtre pour le nom de la fonction
    # logger.addFilter(FunctionNameFilter())

    # Format de base pour les logs
    log_format = (
        '%(asctime)s | %(levelname)-8s | '
        '%(threadName)s | '
        '%(name)s.%(funcName)s:%(lineno)d | '
        '%(message)s'
    )

    # Configuration du fichier de log avec rotation
    log_file = Path(logs_dir) / f"{name}.log"
    file_handler = SafeRotatingFileHandler(
        log_file,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8"
    )
    # Configuration des formateurs
    file_handler.setFormatter(logging.Formatter(log_format, date_format))
    file_handler.setLevel(file_level)
    logger.addHandler(file_handler)

    # Handler console avec couleurs
    console_handler = logging.StreamHandler()
    if enable_colors:
        formatter = ColorFormatter(log_format, date_format)
    else:
        formatter = logging.Formatter(log_format, date_format)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(console_level)
    logger.addHandler(console_handler)

    # Affichage du message d'initialisation du logger dans la console et les logs de debug'
    logger.debug("Logger initialisé avec succès")

    return logger

def get_logger(name: Optional[str] = None) -> logging.Logger:
    """Récupère un logger configuré ou le logger racine"""
    return logging.getLogger(name or __name__)