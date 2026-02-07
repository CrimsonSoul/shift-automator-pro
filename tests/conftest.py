import sys
from unittest.mock import MagicMock

# Mock Windows-specific modules for cross-platform testing.
# Named mocks provide clearer error messages when unexpected attributes
# are accessed (vs. bare MagicMock which silently returns new mocks).
sys.modules["win32print"] = MagicMock(name="win32print")
sys.modules["pythoncom"] = MagicMock(name="pythoncom")
sys.modules["win32com"] = MagicMock(name="win32com")
sys.modules["win32com.client"] = MagicMock(name="win32com.client")
sys.modules["tkcalendar"] = MagicMock(name="tkcalendar")
