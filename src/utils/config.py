"""
Module de gestion de la configuration pour le projet Publipostage DOETH.

Ce module gère le chargement et l'accès à la configuration du projet,
avec la résolution des références et des valeurs environnementales.
"""

import logging
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
        base_dir = Path(__file__).parent.parent.parent
        if config_path is None:
            config_path = str(base_dir / 'config' / 'config.yaml')

        with open(config_path, 'r', encoding='utf-8') as f:
            yaml_content = f.read()

        base_dir_str = str(base_dir).replace('\\', '/')
        yaml_content = yaml_content.replace('__BASE_DIR__', base_dir_str)
        self.config = yaml.safe_load(yaml_content)
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
        """Résout les références ${...} dans la configuration."""
        for key, value in config_dict.items():
            if isinstance(value, dict):
                self._resolve_references(value)
            elif isinstance(value, str):
                matches = re.findall(r'\${([\w.]+)}', value)
                for match in matches:
                    ref_value = self._get_nested_value(match)
                    if ref_value is not None:
                        config_dict[key] = value.replace(f'${{{match}}}', str(ref_value))

    def _get_nested_value(self, path: str) -> Optional[Any]:
        """Récupère une valeur imbriquée à partir d'un chemin comme 'paths.base_dir'."""
        parts = path.split('.')
        current = self.config

        for part in parts:
            if isinstance(current, dict) and part in current:
                current = current[part]
            else:
                return None

        return current

    def get(self, path: str, default: Any = None) -> Any:
        """Récupère une valeur de configuration par son chemin."""
        value = self._get_nested_value(path)
        return value if value is not None else default


# ── Singleton lazy — chargé à la première utilisation ou après init_config() ──
_config: Optional[Config] = None


def init_config(path: Optional[str] = None) -> None:
    """
    Initialise (ou réinitialise) la configuration globale.

    Doit être appelé avant tout get() si un chemin personnalisé est fourni via --config.
    Sans appel explicite, get() initialise automatiquement avec la config par défaut.

    Args:
        path: Chemin vers un fichier de configuration personnalisé, ou None pour la config par défaut.
    """
    global _config
    _config = Config(config_path=path)


def get(path: str, default: Any = None) -> Any:
    """
    Fonction utilitaire pour accéder à la configuration globale.

    Initialise la config par défaut si elle n'a pas encore été chargée.

    Args:
        path: Chemin d'accès à la valeur dans la configuration
        default: Valeur par défaut si le chemin n'existe pas

    Returns:
        Any: Valeur de configuration ou valeur par défaut
    """
    global _config
    if _config is None:
        _config = Config()
    return _config.get(path, default)
