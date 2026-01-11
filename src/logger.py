"""
Logging configuration for Shift Automator application.

This module sets up the logging system with both file and console handlers.
"""

import logging
import os
from pathlib import Path
from typing import Optional

from .constants import LOG_FILENAME


def _get_default_log_dir() -> Path:
    """
    Get the default log directory based on the operating system.

    Returns:
        Path to the default log directory
    """
    if os.name == 'nt':  # Windows
        # Use %LOCALAPPDATA% for Windows
        appdata = os.environ.get('LOCALAPPDATA')
        if appdata:
            return Path(appdata) / 'ShiftAutomator'
        # Fallback to user profile
        user_profile = os.environ.get('USERPROFILE')
        if user_profile:
            return Path(user_profile) / '.shift_automator'
    else:  # macOS/Linux
        # Use ~/.local/share for Linux/macOS
        home = Path.home()
        return home / '.local' / 'share' / 'shift_automator'
    
    # Final fallback to current directory
    return Path('.')


def setup_logging(
    log_level: int = logging.INFO,
    log_dir: Optional[str] = None,
    log_filename: str = LOG_FILENAME
) -> logging.Logger:
    """
    Set up logging configuration for the application.

    Args:
        log_level: The logging level (default: logging.INFO)
        log_dir: Directory for log files (default: AppData directory)
        log_filename: Name of the log file

    Returns:
        Configured logger instance
    """
    # Determine log file path
    if log_dir:
        log_path = Path(log_dir)
    else:
        log_path = _get_default_log_dir()
    
    log_path.mkdir(parents=True, exist_ok=True)
    log_file = log_path / log_filename

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

    # File handler (detailed)
    try:
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
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
