import sys
import os
import pytest
from unittest.mock import MagicMock

# Mock Windows-specific modules before any app code is imported
class MockModule(MagicMock):
    pass

# Setup mocks for Windows-only modules to allow cross-platform testing
if sys.platform != "win32":
    sys.modules["win32print"] = MockModule()
    sys.modules["pythoncom"] = MockModule()
    sys.modules["win32com"] = MockModule()
    sys.modules["win32com.client"] = MockModule()

# Force mock tkinter to avoid display issues in CI/headless
# We create a fake class for Tk so isinstance works if needed
class FakeTk(MagicMock):
    pass

mock_tk = MockModule()
mock_tk.Tk = FakeTk

# Always mock these in non-Windows or CI environments
if sys.platform != "win32" or os.environ.get("GITHUB_ACTIONS") == "true":
    sys.modules["tkinter"] = mock_tk
    sys.modules["tkinter.ttk"] = MockModule()
    sys.modules["tkinter.messagebox"] = MockModule()
    sys.modules["tkinter.filedialog"] = MockModule()
    sys.modules["tkcalendar"] = MockModule()

@pytest.fixture
def mock_wp():
    from .mocks import MockWordProcessor
    return MockWordProcessor()

@pytest.fixture
def mock_ui():
    from .mocks import MockUI
    return MockUI()
