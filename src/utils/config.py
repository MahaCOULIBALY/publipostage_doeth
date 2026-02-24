"""
Module de gestion de la configuration pour le projet Publipostage DOETH.

Ce module gère le chargement et l'accès à la configuration du projet,
avec la résolution des références et des valeurs environnementales.
"""

import logging
import os
import platform
import re
from pathlib import Path
from typing import Optional, Any, Dict, Tuple

import yaml


class Config:
    """Classe de gestion de la configuration"""

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialise la configuration depuis un fichier YAML.

        Args:
            config_path: Chemin vers le fichier de configuration. Si None, utilise le chemin par défaut.
        """
        if config_path is None:
            # Chemin par défaut
            base_dir = Path(__file__).parent.parent.parent
            config_path = os.path.join(base_dir, 'config', 'config.yaml')

        # Charger la configuration
        with open(config_path, 'r', encoding='utf-8') as f:
            yaml_content = f.read()

            # Remplacer __BASE_DIR__ par le chemin absolu
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            yaml_content = yaml_content.replace('__BASE_DIR__', base_dir.replace('\\', '/'))

            # Charger le YAML
            self.config = yaml.safe_load(yaml_content)

        # Résoudre les références de type ${...}
        self._resolve_references(self.config)

    def get_environment(self) -> Tuple[str, Path]:
        """
        Détermine l'environnement d'exécution et le fichier de configuration associé.

        Returns:
            Tuple[str, Path]: Nom de l'environnement et chemin vers le fichier de configuration associé
        """
        hostname = platform.node().upper()
        env_mapping = {
            "FRDC1SRVREBP01": ("PRODUCTION", ".env.prod"),
            "RN-SIEGE787": ("DÉVELOPPEMENT", ".env.develop")
        }
        env_name, env_file = env_mapping.get(hostname, ("TEST", ".env.test"))
        return env_name, Path(__file__).parent.parent.parent / "config" / env_file

    @staticmethod
    def get_log_level(env_name: str) -> int:
        """
        Détermine le niveau de log selon l'environnement.

        Args:
            env_name: Nom de l'environnement (PRODUCTION, DÉVELOPPEMENT, TEST)

        Returns:
            int: Niveau de log correspondant
        """
        return {
            "PRODUCTION": logging.INFO,
            "DÉVELOPPEMENT": logging.DEBUG,
            "TEST": logging.DEBUG
        }.get(env_name, logging.INFO)

    def _resolve_references(self, config_dict: Dict[str, Any]) -> None:
        """
        Résout les références dans la configuration (comme ${paths.base_dir}).

        Args:
            config_dict: Dictionnaire de configuration à traiter
        """
        for key, value in config_dict.items():
            if isinstance(value, dict):
                self._resolve_references(value)
            elif isinstance(value, str):
                # Rechercher et remplacer les références ${...}
                matches = re.findall(r'\${([\w.]+)}', value)
                for match in matches:
                    # Obtenir la valeur référencée
                    ref_value = self._get_nested_value(match)
                    if ref_value is not None:
                        # Remplacer la référence par sa valeur
                        config_dict[key] = value.replace(f'${{{match}}}', str(ref_value))

    def _get_nested_value(self, path: str) -> Optional[Any]:
        """
        Récupère une valeur imbriquée à partir d'un chemin comme 'paths.base_dir'.

        Args:
            path: Chemin d'accès à la valeur dans la configuration

        Returns:
            Any: Valeur trouvée ou None si le chemin n'existe pas
        """
        parts = path.split('.')
        current = self.config

        for part in parts:
            if isinstance(current, dict) and part in current:
                current = current[part]
            else:
                return None

        return current

    def get(self, path: str, default: Any = None) -> Any:
        """
        Récupère une valeur de configuration par son chemin.

        Args:
            path: Chemin d'accès à la valeur dans la configuration
            default: Valeur par défaut si le chemin n'existe pas

        Returns:
            Any: Valeur de configuration ou valeur par défaut
        """
        value = self._get_nested_value(path)
        return value if value is not None else default


# Instance globale de la configuration
config = Config()


# Fonction d'accès pour importer directement des valeurs
def get(path: str, default: Any = None) -> Any:
    """
    Fonction utilitaire pour accéder à la configuration globale.

    Args:
        path: Chemin d'accès à la valeur dans la configuration
        default: Valeur par défaut si le chemin n'existe pas

    Returns:
        Any: Valeur de configuration ou valeur par défaut
    """
    return config.get(path, default)