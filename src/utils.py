"""
Shared utility functions for Shift Automator application.

This module contains common utility functions used across multiple modules.
"""

import os
from pathlib import Path
from typing import Optional, List, Tuple

try:
    import tkinter.font as tkfont
    HAS_TKINTER = True
except ImportError:
    HAS_TKINTER = False


def get_app_data_dir(app_name: str = "ShiftAutomator") -> Path:
    """
    Get the application data directory based on the operating system.

    Args:
        app_name: Name of the application directory

    Returns:
        Path to the application data directory
    """
    if os.name == 'nt':  # Windows
        # Use %LOCALAPPDATA% for Windows
        appdata = os.environ.get('LOCALAPPDATA')
        if appdata:
            return Path(appdata) / app_name
        # Fallback to user profile
        user_profile = os.environ.get('USERPROFILE')
        if user_profile:
            return Path(user_profile) / f'.{app_name.lower()}'
    else:  # macOS/Linux
        # Use ~/.local/share for Linux/macOS
        home = Path.home()
        return home / '.local' / 'share' / app_name.lower()
    
    # Final fallback to current directory
    return Path('.')


def get_available_font(preferred_fonts: List[str], size: int, weight: str = "normal") -> Tuple:
    """
    Return the first available font from a list, or a fallback.

    This helps provide cross-platform font support by falling back through
    a list of preferred fonts to a system default.

    Args:
        preferred_fonts: List of font names to try, in order of preference
        size: Font size in points
        weight: Font weight - "normal" or "bold"

    Returns:
        Font tuple compatible with Tkinter (font_name, size, [weight])

    Examples:
        >>> font = get_available_font(["Segoe UI", "Arial"], 10, "bold")
        >>> font
        ('Segoe UI', 10, 'bold')
    """
    if not HAS_TKINTER:
        # Fallback for environments without tkinter (e.g., tests)
        if weight == "bold":
            return ("TkDefaultFont", size, "bold")
        return ("TkDefaultFont", size)

    try:
        available = set(tkfont.families())
        for font_name in preferred_fonts:
            if font_name in available:
                if weight == "bold":
                    return (font_name, size, "bold")
                return (font_name, size)
    except Exception:
        pass

    # Ultimate fallback to system default
    if weight == "bold":
        return ("TkDefaultFont", size, "bold")
    return ("TkDefaultFont", size)
