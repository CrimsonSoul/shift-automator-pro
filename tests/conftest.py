import sys
from unittest.mock import MagicMock

# Mock Windows-specific modules
class MockModule(MagicMock):
    pass

sys.modules["win32print"] = MockModule()
sys.modules["pythoncom"] = MockModule()
sys.modules["win32com"] = MockModule()
sys.modules["win32com.client"] = MockModule()
sys.modules["tkcalendar"] = MockModule()
