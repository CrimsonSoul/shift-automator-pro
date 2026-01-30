"""
Logging configuration for Shift Automator application.

This module sets up the logging system with both file and console handlers.
"""

import logging
import logging.handlers
import os
from pathlib import Path
from typing import Optional

from .constants import LOG_FILENAME

__all__ = ["setup_logging", "get_logger"]


def setup_logging(
    log_level: Optional[int] = None,
    log_dir: Optional[str] = None,
    log_filename: str = LOG_FILENAME
) -> logging.Logger:
    """
    Set up logging configuration for the application.

    Args:
        log_level: The logging level (default: check DEBUG env or logging.INFO)
        log_dir: Directory for log files (default: current working directory)
        log_filename: Name of the log file

    Returns:
        Configured logger instance
    """
    if log_level is None:
        if os.environ.get("DEBUG") == "1" or os.environ.get("DEBUG", "").lower() == "true":
            log_level = logging.DEBUG
        else:
            log_level = logging.INFO

    # Determine log file path
    if log_dir:
        log_path = Path(log_dir)
        log_path.mkdir(parents=True, exist_ok=True)
        log_file = log_path / log_filename
    else:
        log_file = Path(log_filename)

    # Create logger
    logger = logging.getLogger("shift_automator")
    logger.setLevel(log_level)

    # Remove existing handlers to avoid duplicates
    logger.handlers.clear()

    # Create formatters
    detailed_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    simple_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    )

    # File handler (detailed, with rotation: 5MB max, keep 3 backups)
    try:
        file_handler = logging.handlers.RotatingFileHandler(
            log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(detailed_formatter)
        logger.addHandler(file_handler)
    except (IOError, OSError) as e:
        # If we can't write to file, at least log to console
        print(f"Warning: Could not create log file: {e}")

    # Console handler (simple)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(simple_formatter)
    logger.addHandler(console_handler)

    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """
    Get a logger instance.

    Args:
        name: Logger name (default: "shift_automator")

    Returns:
        Logger instance
    """
    return logging.getLogger(name or "shift_automator")
