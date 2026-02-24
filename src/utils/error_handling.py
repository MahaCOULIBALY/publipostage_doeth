# src/utils/error_handling.py
import logging
import traceback
from functools import wraps
from typing import Callable, Any, Optional, Dict
from datetime import datetime


def handle_errors(func: Optional[Callable] = None) -> Callable:
    """
    Decorator for standardized error handling in ETL functions.

    Args:
        func: The function to decorate

    Returns:
        Callable: Decorated function with error handling
    """

    def _get_logger() -> logging.Logger:
        return logging.getLogger('etl-pipeline')

    def decorator(fn: Callable) -> Callable:
        logger = _get_logger()

        @wraps(fn)
        def wrapper(*args, **kwargs) -> Any:
            try:
                return fn(*args, **kwargs)
            except Exception as e:
                logger.error(
                    f"Error in {fn.__name__}: {str(e)}",
                    exc_info=True
                )
                raise

        return wrapper

    if func is not None:
        return decorator(func)
    return decorator


class ETLError(Exception):
    """Base exception class for all ETL pipeline errors."""

    def __init__(self, message: str, original_error: Optional[Exception] = None):
        super().__init__(message)
        self.original_error = original_error
        self.timestamp = datetime.now()
        self.traceback = traceback.format_exc() if original_error else None


class ValidationError(ETLError):
    """Exception for data validation errors in the ETL pipeline.

    Examples:
        - Missing required fields
        - Invalid data types
        - Business rule violations
        - Schema validation failures
    """

    def __init__(self, message: str, field: Optional[str] = None,
                 value: Optional[Any] = None, original_error: Optional[Exception] = None):
        super().__init__(message, original_error)
        self.field = field
        self.invalid_value = value


class ConfigurationError(ETLError):
    """Exception for configuration-related errors in the ETL pipeline.

    Examples:
        - Missing configuration parameters
        - Invalid configuration values
        - Missing environment variables
        - Invalid file paths
    """

    def __init__(self, message: str, config_key: Optional[str] = None,
                 original_error: Optional[Exception] = None):
        super().__init__(message, original_error)
        self.config_key = config_key


class DatabaseError(Exception):
    """Custom exception for database-related errors in the ETL pipeline."""

    def __init__(self, message: str, original_error: Exception = None):
        super().__init__(message)
        self.original_error = original_error


class ProcessingError(Exception):
    """
    Exception spécifique pour les erreurs de traitement dans le pipeline ETL.
    Utilisée pour signaler les problèmes lors de l'exécution du pipeline.
    """

    ERROR_CODES = {
        'VALIDATION_FAILED': 'Échec de la validation des données',
        'TRANSFORMATION_FAILED': 'Échec de la transformation des données',
        'LOADING_FAILED': 'Échec du chargement des données',
        'EXECUTION_FAILED': "Échec de l'exécution du pipeline",
        'CONFIG_ERROR': 'Erreur de configuration',
        'RESOURCE_ERROR': 'Erreur de ressource'
    }

    def __init__(self, message: str, error_type: str = 'EXECUTION_FAILED',
                 context: Optional[Dict[str, Any]] = None,
                 source: Optional[str] = None) -> None:
        """
        Initialise une erreur de traitement avec contexte enrichi.

        Args:
            message: Description de l'erreur
            error_type: Type d'erreur parmi les ERROR_CODES définis
            context: Informations contextuelles sur l'erreur
            source: Source ou composant à l'origine de l'erreur
        """
        if error_type not in self.ERROR_CODES:
            error_type = 'EXECUTION_FAILED'

        details = {
            'error_type': error_type,
            'error_description': self.ERROR_CODES[error_type],
            'source': source or 'unknown'
        }

        if context:
            details.update(context)

        super().__init__(
            message=message,
            error_code=f"ETL_{error_type}",
            details=details
        )
