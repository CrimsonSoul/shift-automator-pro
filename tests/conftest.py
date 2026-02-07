import sys
from unittest.mock import MagicMock

# Mock Windows-specific modules
sys.modules["win32print"] = MagicMock()
sys.modules["pythoncom"] = MagicMock()
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()
sys.modules["tkcalendar"] = MagicMock()
