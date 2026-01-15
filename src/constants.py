"""
Constants for Shift Automator application.

This module contains all named constants used throughout the application
to avoid magic numbers and strings.
"""

from dataclasses import dataclass
from typing import Final

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

# Replacement Tokens
DATE_PLACEHOLDER: Final = "{{DATE}}"

# UI placeholder values
PRINTER_DEFAULT_PLACEHOLDER: Final = "Choose Printer"

# UI Constants
WINDOW_WIDTH: Final = 640
WINDOW_HEIGHT: Final = 880
WINDOW_RESIZABLE: Final = False

# Progress bar
PROGRESS_MAX: Final = 100

# Retry settings for COM calls
COM_RETRIES: Final = 5
COM_RETRY_DELAY: Final = 1  # seconds
COM_TIMEOUT: Final = 30  # seconds - timeout for COM operations

# COM Initialization Settings
COM_INIT_MAX_RETRIES_PER_METHOD: Final = 2  # Max retries per dispatch method after cache clear
COM_INIT_CACHE_CLEAR_DELAY: Final = 2.0     # Seconds to wait after clearing cache
COM_INIT_PROCESS_KILL_DELAY: Final = 2.0    # Seconds to wait after killing Word processes
COM_INIT_STABILIZATION_DELAY: Final = 3.0   # Seconds to wait after successful dispatch (Word needs time)
COM_INIT_VERIFICATION_RETRIES: Final = 5    # Number of times to retry verification
COM_INIT_VERIFICATION_DELAY: Final = 1.0    # Seconds between verification retries

# COM Error Codes (hexadecimal)
COM_ERROR_RPC_CALL_REJECTED: Final = "0x80010001"  # RPC_E_CALL_REJECTED
COM_ERROR_RPC_SERVERCALL_RETRYLATER: Final = "0x80010101"  # RPC_E_SERVERCALL_RETRYLATER
COM_ERROR_DISP_E_EXCEPTION: Final = "0x80020009"  # DISP_E_EXCEPTION (-2147352567)
COM_ERROR_DISP_E_BADINDEX: Final = "0x8002000b"   # DISP_E_BADINDEX (-2147352565)

# Printer Status Flags (from Windows API)
PRINTER_STATUS_OFFLINE: Final = 0x00000080
PRINTER_STATUS_ERROR: Final = 0x00000002

# Printer check timeout
PRINTER_STATUS_TIMEOUT: Final = 5.0  # seconds

# Retry settings for print operations
PRINT_MAX_RETRIES: Final = 3
PRINT_INITIAL_DELAY: Final = 2.0  # seconds
PRINT_MAX_DELAY: Final = 10.0  # seconds

# Transient error keywords for retry
TRANSIENT_ERROR_KEYWORDS: Final = (
    "offline",
    "not ready",
    "busy",
    "timeout",
    "temporarily",
    "unavailable"
)

# Configuration save debouncing
CONFIG_DEBOUNCE_DELAY: Final = 1.0  # seconds - delay before saving config


# Printer status check caching
PRINTER_STATUS_CACHE_TTL: Final = 30  # seconds


@dataclass(frozen=True)
class Colors:
    """Color scheme constants for the application UI (Relay Design System)."""
    background: str = "#0B0D12"      # Relay app background
    surface: str = "#18181B"         # Relay surface/panel
    surface_elevated: str = "#1E1E21" # Relay elevated surface
    accent: str = "#3B82F6"          # Relay accent blue
    text_main: str = "#FFFFFF"       # Relay primary text
    text_dim: str = "#A1A1AA"        # Relay secondary text
    text_tertiary: str = "#71717A"   # Relay tertiary text
    success: str = "#10B981"         # Relay success green
    error: str = "#FF5C5C"           # Relay danger
    border: str = "#27272A"          # Relay border (approximate for solid)
    secondary: str = "#27272A"       # Button hover / secondary action
    accent_hover: str = "#60A5FA"    # Relay accent blue hover


@dataclass(frozen=True)
class Fonts:
    """
    Font configuration for the application UI.

    Font resolution happens at runtime to support cross-platform compatibility.
    See utils.get_available_font() for resolution logic.
    """
    main: tuple = ("Segoe UI", 10)
    bold: tuple = ("Segoe UI", 10, "bold")
    header: tuple = ("Segoe UI", 24, "bold")
    sub: tuple = ("Segoe UI", 9)
    button: tuple = ("Segoe UI", 11, "bold")


# Font preference list for runtime resolution
# These will be tried in order when fonts are initialized in UI
FONT_PREFERENCES: Final = [
    "Segoe UI Variable Display",
    "Segoe UI",
    "Arial",
    "Helvetica",
    "Tahoma"
]


# Global color and font instances
COLORS = Colors()
FONTS = Fonts()
