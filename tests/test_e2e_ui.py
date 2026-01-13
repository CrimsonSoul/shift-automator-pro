import pytest
import tkinter as tk
from unittest.mock import MagicMock, patch
from src import ScheduleAppUI, ShiftAutomatorApp
from .mocks import MockWordProcessor

@pytest.fixture
def root():
    # Since tkinter is mocked in conftest.py, tk.Tk() returns a Mock object.
    # We just need to ensure it has the methods the app calls.
    root = tk.Tk()
    return root

def test_ui_initialization(root):
    """Test that UI components are created correctly."""
    ui = ScheduleAppUI(root)
    # Check if title was called on root
    root.title.assert_called_with("Shift Automator")
        
    assert ui.day_entry is not None
    assert ui.night_entry is not None
    assert ui.print_btn is not None

def test_ui_validation_triggers(root):
    """Test that the 'Start' button state changes based on input."""
    ui = ScheduleAppUI(root)
    app = ShiftAutomatorApp(root)
    app.ui = ui
    
    # Mock the show_warning method on the instance
    ui.show_warning = MagicMock()
    
    # Simulate invalid inputs by mocking the UI's get methods
    with patch.object(ui, "get_day_folder", return_value=""):
        app.start_processing()
    
    # Check if show_warning was called
    ui.show_warning.assert_called()
    args, kwargs = ui.show_warning.call_args
    assert "Validation Error" in args[0]

def test_ui_start_execution_flow(root):
    """Test that clicking start triggers the expected methods."""
    ui = ScheduleAppUI(root)
    app = ShiftAutomatorApp(root)
    app.ui = ui
    
    # Mock validation and thread start
    with patch.object(app, "_validate_inputs", return_value=(True, None)):
        with patch("threading.Thread") as mock_thread:
            app.start_processing()
            mock_thread.assert_called_once()
            kwargs = mock_thread.call_args.kwargs
            assert kwargs["target"] == app._process_batch

def test_ui_cancel_flow(root):
    """Test that the cancel button works."""
    ui = ScheduleAppUI(root)
    app = ShiftAutomatorApp(root)
    app.ui = ui
    
    # Mock the update_status method on the instance
    ui.update_status = MagicMock()
    
    app._processing_thread = MagicMock()
    app._processing_thread.is_alive.return_value = True
    
    app.cancel_processing()
    
    assert app._cancel_requested is True
    ui.update_status.assert_called()
    args, _ = ui.update_status.call_args
    assert "Cancellation requested" in args[0]
