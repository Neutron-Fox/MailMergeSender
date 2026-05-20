"""
Logging configuration module for Universal Email Sender application.
Provides centralized logging setup with file rotation and console output.
"""
import logging
import os
import sys
import io
from logging.handlers import RotatingFileHandler

from .constants import (
    LOG_DIR,
    LOG_FILE,
    LOG_LEVEL,
    LOG_MAX_SIZE,
    LOG_BACKUP_COUNT,
    LOG_FORMAT,
    LOG_DATE_FORMAT
)


class LoggerSetup:
    """Centralizes logging configuration"""
    
    _loggers = {}
    
    @staticmethod
    def setup_logging():
        """Configure root logger and file handler"""
        # Ensure log directory exists
        os.makedirs(LOG_DIR, exist_ok=True)
        
        # Create root logger
        root_logger = logging.getLogger()
        root_logger.setLevel(getattr(logging, LOG_LEVEL))
        
        # Remove existing handlers
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        
        # Create formatter
        formatter = logging.Formatter(LOG_FORMAT, datefmt=LOG_DATE_FORMAT)
        
        # File handler with rotation
        file_handler = RotatingFileHandler(
            LOG_FILE,
            maxBytes=LOG_MAX_SIZE,
            backupCount=LOG_BACKUP_COUNT
        )
        file_handler.setLevel(getattr(logging, LOG_LEVEL))
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
        
        # Console handler with UTF-8 encoding to support unicode characters
        # Wrap stdout with UTF-8 encoding to handle unicode characters like ✓
        try:
            if sys.stdout.encoding.lower() != 'utf-8':
                # Reconfigure stdout to use UTF-8 encoding with error handling
                sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        except (AttributeError, TypeError):
            pass
        
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(getattr(logging, LOG_LEVEL))
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)
        
        logging.info(f"Logging initialized - Level: {LOG_LEVEL}, File: {LOG_FILE}")
    
    @staticmethod
    def get_logger(name: str) -> logging.Logger:
        """
        Get logger for module.
        
        Args:
            name: Logger name (usually __name__)
            
        Returns:
            Configured logger instance
        """
        if name not in LoggerSetup._loggers:
            logger = logging.getLogger(name)
            LoggerSetup._loggers[name] = logger
        
        return LoggerSetup._loggers[name]


def get_logger(name: str) -> logging.Logger:
    """Convenience function to get logger"""
    return LoggerSetup.get_logger(name)
