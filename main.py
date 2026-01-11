"""
Shift Automator - Main Entry Point

This is the main entry point for the Shift Automator application.
All application logic has been refactored into the src package.
"""

import sys
from pathlib import Path

# Add src directory to path for imports
src_path = Path(__file__).parent / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from src.main import main

if __name__ == "__main__":
    main()
