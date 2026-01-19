"""
Constants for Shift Automator application.

This module contains all named constants used throughout the application
to avoid magic numbers and strings.
"""

from dataclasses import dataclass
from typing import Final

__all__ = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY",
           "PROTECTION_NONE", "PROTECTION_READ_ONLY", "PROTECTION_ALLOW_COMMENTS", "PROTECTION_ALLOW_REVISIONS",
           "CLOSE_NO_SAVE", "CLOSE_SAVE", "CLOSE_PROMPT",
           "PRINTER_ENUM_LOCAL", "PRINTER_ENUM_NETWORK",
           "DOCX_EXTENSION", "CONFIG_FILENAME", "LOG_FILENAME",
           "WINDOW_WIDTH", "WINDOW_HEIGHT", "WINDOW_RESIZABLE",
           "PROGRESS_MAX", "MAX_DAYS_RANGE",
           "COM_RETRIES", "COM_RETRY_DELAY",
           "COLORS", "FONTS"]

# Weekday constants (Python's datetime.weekday() returns 0=Monday, 6=Sunday)
MONDAY: Final = 0
TUESDAY: Final = 1
WEDNESDAY: Final = 2
THURSDAY: Final = 3
FRIDAY: Final = 4
SATURDAY: Final = 5
SUNDAY: Final = 6

# Word document protection types
PROTECTION_NONE: Final = -1
PROTECTION_READ_ONLY: Final = 0
PROTECTION_ALLOW_COMMENTS: Final = 1
PROTECTION_ALLOW_REVISIONS: Final = 2

# Word document close options
CLOSE_NO_SAVE: Final = 0
CLOSE_SAVE: Final = 1
CLOSE_PROMPT: Final = 2

# Windows printer enumeration constants
PRINTER_ENUM_LOCAL: Final = 2
PRINTER_ENUM_NETWORK: Final = 4

# File extensions
DOCX_EXTENSION: Final = ".docx"

# Configuration
CONFIG_FILENAME: Final = "config.json"
LOG_FILENAME: Final = "shift_automator.log"

# UI Constants
WINDOW_WIDTH: Final = 640
WINDOW_HEIGHT: Final = 720
WINDOW_RESIZABLE: Final = False

# Progress bar
PROGRESS_MAX: Final = 100

# Date validation
MAX_DAYS_RANGE: Final = 365

# Retry settings for COM calls
COM_RETRIES: Final = 5
COM_RETRY_DELAY: Final = 1  # seconds


@dataclass(frozen=True)
class Colors:
    """Color scheme constants for the application UI."""
    background: str = "#0D0D12"      # Near-black depth
    surface: str = "#16161D"         # Modern card surface
    accent: str = "#4D7CFF"          # Tech blue
    text_main: str = "#FFFFFF"       # High contrast
    text_dim: str = "#71717A"        # Muted secondary
    success: str = "#10B981"         # Emerald
    error: str = "#EF4444"           # Red/Danger
    border: str = "#27272A"          # Subtle border
    secondary: str = "#1E1E26"       # Button hover
    accent_hover: str = "#3A6DFF"    # Accent hover state


@dataclass(frozen=True)
class Fonts:
    """Font configuration for the application UI."""
    main: tuple = ("Segoe UI Variable Display", 10)
    bold: tuple = ("Segoe UI Variable Display", 10, "bold")
    header: tuple = ("Segoe UI Variable Display", 24, "bold")
    sub: tuple = ("Segoe UI Variable Display", 9)
    button: tuple = ("Segoe UI Variable Display", 11, "bold")


# Global color and font instances
COLORS = Colors()
FONTS = Fonts()
