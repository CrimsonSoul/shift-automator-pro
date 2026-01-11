"""
Shift Automator - A high-performance desktop application for automating shift schedule printing.

This package provides modules for:
- Configuration management (config)
- UI components (ui)
- Word document processing (word_processor)
- Date and scheduling logic (scheduler)
- Path validation (path_validation)
- Constants and styling (constants)
- Logging setup (logger)
"""

from .main import ShiftAutomatorApp, main

__version__ = "2.0.0"
__author__ = "Shift Automator Team"

__all__ = ["ShiftAutomatorApp", "main"]
