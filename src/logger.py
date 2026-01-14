"""
Logging configuration for Shift Automator application.

This module sets up the logging system with both file and console handlers.
"""

import logging
import os
import sys
from pathlib import Path
from typing import Optional, List

from .constants import LOG_FILENAME
from .utils import get_app_data_dir


def _get_exe_directory() -> Optional[Path]:
    """
    Get the directory where the executable is located (for PyInstaller bundles).
    
    Returns:
        Path to exe directory, or None if not applicable
    """
    try:
        # PyInstaller sets sys.frozen when running as a bundle
        if getattr(sys, 'frozen', False):
            return Path(sys.executable).parent
    except Exception:
        pass
    return None


def _get_default_log_dir() -> Path:
    """
    Get the default log directory based on the operating system.

    Returns:
        Path to the default log directory
    """
    return get_app_data_dir("ShiftAutomator")


def setup_logging(
    log_level: int = logging.DEBUG,  # Default to DEBUG for better diagnostics
    log_dir: Optional[str] = None,
    log_filename: str = LOG_FILENAME
) -> logging.Logger:
    """
    Set up logging configuration for the application.
    
    Writes logs to multiple locations for easier discovery:
    1. AppData directory (standard location)
    2. Exe directory (for PyInstaller bundles - easier for users to find)

    Args:
        log_level: The logging level (default: logging.DEBUG)
        log_dir: Directory for log files (default: AppData directory)
        log_filename: Name of the log file

    Returns:
        Configured logger instance
    """
    # Create logger
    logger = logging.getLogger("shift_automator")
    logger.setLevel(logging.DEBUG)  # Always capture DEBUG, handlers filter

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

    # Collect all log file paths we want to write to
    log_paths: List[Path] = []
    
    # 1. Primary: AppData directory (or custom log_dir)
    if log_dir:
        log_paths.append(Path(log_dir))
    else:
        log_paths.append(_get_default_log_dir())
    
    # 2. Secondary: Exe directory (for PyInstaller bundles)
    exe_dir = _get_exe_directory()
    if exe_dir and exe_dir not in log_paths:
        log_paths.append(exe_dir)
    
    # Create file handlers for each location
    log_files_created: List[str] = []
    for log_path in log_paths:
        try:
            log_path.mkdir(parents=True, exist_ok=True)
            log_file = log_path / log_filename
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(detailed_formatter)
            logger.addHandler(file_handler)
            log_files_created.append(str(log_file))
        except (IOError, OSError) as e:
            # If we can't write to this location, continue to next
            print(f"Warning: Could not create log file at {log_path}: {e}")

    # Console handler (simple)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(simple_formatter)
    logger.addHandler(console_handler)

    # Log startup banner with file locations
    logger.info("=" * 60)
    logger.info("Shift Automator - Logging Initialized")
    logger.info(f"Python version: {sys.version}")
    logger.info(f"Platform: {sys.platform}")
    logger.info(f"Frozen (PyInstaller): {getattr(sys, 'frozen', False)}")
    if log_files_created:
        logger.info(f"Log files: {', '.join(log_files_created)}")
    else:
        logger.warning("No log files created - logging to console only")
    logger.info("=" * 60)

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
