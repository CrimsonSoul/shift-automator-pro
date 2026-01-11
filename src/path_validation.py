"""
Path validation and safety utilities for Shift Automator application.

This module provides functions to validate and sanitize file paths to prevent
security issues like path traversal attacks.
"""

import os
from pathlib import Path
from typing import Optional, Tuple

from .logger import get_logger

logger = get_logger(__name__)


def validate_folder_path(path: str) -> Tuple[bool, Optional[str]]:
    """
    Validate that a folder path is safe and accessible.

    Args:
        path: The folder path to validate

    Returns:
        Tuple of (is_valid, error_message)

    Security considerations:
        - Resolves relative paths to absolute paths
        - Checks that path exists and is a directory
        - Checks that path is readable
    """
    if not path:
        return False, "Path cannot be empty"

    try:
        # Convert to absolute path to resolve relative paths
        abs_path = Path(path).resolve(strict=True)

        # Check if it's a directory
        if not abs_path.is_dir():
            return False, f"Path is not a directory: {path}"

        # Check if directory is readable
        if not os.access(abs_path, os.R_OK):
            return False, f"Directory is not readable: {path}"

        logger.debug(f"Validated folder path: {abs_path}")
        return True, None

    except FileNotFoundError:
        return False, f"Path does not exist: {path}"
    except PermissionError:
        return False, f"Permission denied accessing path: {path}"
    except OSError as e:
        return False, f"Error accessing path: {e}"


def validate_file_path(path: str, allowed_extensions: Optional[list[str]] = None) -> Tuple[bool, Optional[str]]:
    """
    Validate that a file path is safe and has an allowed extension.

    Args:
        path: The file path to validate
        allowed_extensions: List of allowed file extensions (e.g., [".docx"])

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not path:
        return False, "Path cannot be empty"

    try:
        # Convert to absolute path
        abs_path = Path(path).resolve(strict=True)

        # Check if it's a file
        if not abs_path.is_file():
            return False, f"Path is not a file: {path}"

        # Check file extension if specified
        if allowed_extensions:
            ext = abs_path.suffix.lower()
            if ext not in [e.lower() for e in allowed_extensions]:
                return False, f"File extension '{ext}' not allowed. Allowed: {allowed_extensions}"

        # Check if file is readable
        if not os.access(abs_path, os.R_OK):
            return False, f"File is not readable: {path}"

        logger.debug(f"Validated file path: {abs_path}")
        return True, None

    except FileNotFoundError:
        return False, f"File does not exist: {path}"
    except PermissionError:
        return False, f"Permission denied accessing file: {path}"
    except OSError as e:
        return False, f"Error accessing file: {e}"


def sanitize_filename(filename: str) -> str:
    """
    Sanitize a filename by removing potentially dangerous characters.

    Args:
        filename: The filename to sanitize

    Returns:
        Sanitized filename
    """
    # Remove path separators and other dangerous characters
    dangerous_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    sanitized = filename
    for char in dangerous_chars:
        sanitized = sanitized.replace(char, '_')

    # Remove leading/trailing dots and spaces
    sanitized = sanitized.strip('. ')

    # Limit length
    if len(sanitized) > 255:
        sanitized = sanitized[:255]

    logger.debug(f"Sanitized filename: '{filename}' -> '{sanitized}'")
    return sanitized


def is_path_within_base(path: str, base_path: str) -> bool:
    """
    Check if a path is within a base directory (prevents directory traversal).

    Args:
        path: The path to check
        base_path: The base directory path

    Returns:
        True if path is within base_path, False otherwise
    """
    try:
        abs_path = Path(path).resolve()
        abs_base = Path(base_path).resolve()

        # Check if abs_path is a subdirectory of abs_base
        try:
            abs_path.relative_to(abs_base)
            return True
        except ValueError:
            return False
    except (OSError, ValueError):
        return False
