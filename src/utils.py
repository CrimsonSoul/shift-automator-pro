"""
Shared utility functions for Shift Automator application.

This module contains common utility functions used across multiple modules.
"""

import os
from pathlib import Path
from typing import Optional


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
