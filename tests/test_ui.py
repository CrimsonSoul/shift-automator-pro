"""
Unit tests for UI components module.

These tests verify the UI component initialization, styling, and interaction handlers.
"""

import sys
import os
import tkinter as tk
from unittest.mock import MagicMock, patch, Mock
import pytest

from src.ui import ScheduleAppUI, HAS_WIN32PRINT
from src.constants import COLORS, FONTS

# Skip UI tests on platforms without display support or in CI
needs_display = pytest.mark.skipif(
    sys.platform == "darwin" or sys.platform.startswith("linux") or os.environ.get("GITHUB_ACTIONS") == "true",
    reason="UI tests require display server (skipped in CI or non-Windows)"
)


@pytest.fixture
def root():
    """Create a test Tkinter root window."""
    try:
        root = tk.Tk()
        root.withdraw()  # Hide the window
        yield root
        root.destroy()
    except Exception:
        # If Tkinter fails to initialize (e.g. headless), yield a mock
        # This allows the rest of the test suite to load even if this file is skipped
        root = MagicMock()
        yield root


@needs_display
class TestScheduleAppUI:
    """Unit tests for ScheduleAppUI class."""

    def test_initialization(self, root):
        """Test UI initialization."""
        ui = ScheduleAppUI(root)

        assert ui.root == root
        assert ui.day_entry is not None
        assert ui.night_entry is not None
        assert ui.start_date_picker is not None
        assert ui.end_date_picker is not None
        assert ui.printer_var is not None
        assert ui.log_widget is not None
        assert ui.progress is not None
        assert ui.print_btn is not None
        assert ui.cancel_btn is not None

    def test_window_configuration(self, root):
        """Test window is configured correctly."""
        ui = ScheduleAppUI(root)

        assert "Shift Automator" in str(root.title())
        # Check geometry
        geometry = str(root.geometry())
        assert "640" in geometry
        assert "720" in geometry

    def test_get_day_folder(self, root):
        """Test getting day folder path."""
        ui = ScheduleAppUI(root)
        test_path = "/path/to/day/folder"

        ui.day_entry.delete(0, tk.END)
        ui.day_entry.insert(0, test_path)

        result = ui.get_day_folder()
        assert result == test_path

    def test_get_day_folder_empty(self, root):
        """Test getting empty day folder."""
        ui = ScheduleAppUI(root)

        result = ui.get_day_folder()
        assert result == ""

    def test_get_night_folder(self, root):
        """Test getting night folder path."""
        ui = ScheduleAppUI(root)
        test_path = "/path/to/night/folder"

        ui.night_entry.delete(0, tk.END)
        ui.night_entry.insert(0, test_path)

        result = ui.get_night_folder()
        assert result == test_path

    def test_get_printer_name(self, root):
        """Test getting printer name."""
        ui = ScheduleAppUI(root)
        test_printer = "Test Printer"

        ui.printer_var.set(test_printer)

        result = ui.get_printer_name()
        assert result == test_printer

    def test_get_start_date(self, root):
        """Test getting start date."""
        ui = ScheduleAppUI(root)

        result = ui.get_start_date()
        assert result is not None

    def test_get_end_date(self, root):
        """Test getting end date."""
        ui = ScheduleAppUI(root)

        result = ui.get_end_date()
        assert result is not None

    def test_set_start_command(self, root):
        """Test setting start button command."""
        ui = ScheduleAppUI(root)
        mock_command = MagicMock()

        ui.set_start_command(mock_command)

        # Trigger button click
        ui.print_btn.invoke()
        mock_command.assert_called_once()

    def test_set_cancel_command(self, root):
        """Test setting cancel button command."""
        ui = ScheduleAppUI(root)
        mock_command = MagicMock()

        ui.set_cancel_command(mock_command)

        # Cancel button starts disabled, so enable it first
        ui.cancel_btn.config(state="normal")
        ui.cancel_btn.invoke()
        mock_command.assert_called_once()

    def test_set_print_button_state(self, root):
        """Test setting print button state."""
        ui = ScheduleAppUI(root)

        ui.set_print_button_state("disabled")
        assert ui.print_btn.cget("state") == "disabled"

        ui.set_print_button_state("normal")
        assert ui.print_btn.cget("state") == "normal"

    def test_set_cancel_button_state(self, root):
        """Test setting cancel button state."""
        ui = ScheduleAppUI(root)

        ui.set_cancel_button_state("disabled")
        assert ui.cancel_btn.cget("state") == "disabled"

        ui.set_cancel_button_state("normal")
        assert ui.cancel_btn.cget("state") == "normal"

    def test_update_progress(self, root):
        """Test updating progress bar."""
        ui = ScheduleAppUI(root)

        ui.update_progress(50.0)
        assert ui.progress_var.get() == 50.0

    def test_log(self, root):
        """Test appending message to log widget."""
        ui = ScheduleAppUI(root)
        test_message = "Test activity log message"
        
        ui.log(test_message)
        
        # In read-only mode it's hard to check content directly without enabling
        # but we can check if it was inserted
        content = ui.log_widget.get("1.0", tk.END)
        assert test_message in content

    @patch('src.ui.messagebox.showerror')
    def test_show_error(self, mock_showerror, root):
        """Test showing error message."""
        ui = ScheduleAppUI(root)

        ui.show_error("Error Title", "Error message")

        mock_showerror.assert_called_once_with("Error Title", "Error message")

    @patch('src.ui.messagebox.showwarning')
    def test_show_warning(self, mock_showwarning, root):
        """Test showing warning message."""
        ui = ScheduleAppUI(root)

        ui.show_warning("Warning Title", "Warning message")

        mock_showwarning.assert_called_once_with("Warning Title", "Warning message")

    @patch('src.ui.messagebox.showinfo')
    def test_show_info(self, mock_showinfo, root):
        """Test showing info message."""
        ui = ScheduleAppUI(root)

        ui.show_info("Info Title", "Info message")

        mock_showinfo.assert_called_once_with("Info Title", "Info message")

    def test_config_change_callback(self, root):
        """Test configuration change callback."""
        ui = ScheduleAppUI(root)
        mock_callback = MagicMock()

        ui.set_config_change_callback(mock_callback)

        # Trigger callback by changing printer
        ui.printer_var.set("New Printer")

        mock_callback.assert_called_once()

    def test_printer_dropdown_without_win32print(self, root):
        """Test printer dropdown when win32print is not available."""
        with patch('src.ui.HAS_WIN32PRINT', False):
            ui = ScheduleAppUI(root)

            # Should show error message
            assert ui.printer_var is not None
            # In mock environment, it might be empty or "Not Available" depending on state
            assert isinstance(ui.printer_var.get(), str)

    def test_style_configuration(self, root):
        """Test UI styles are configured."""
        ui = ScheduleAppUI(root)

        # Check style was created
        assert ui.style is not None

        # Check theme was set
        assert str(ui.style.theme_use()) == 'clam'

    def test_color_constants(self):
        """Test color constants are defined."""
        assert hasattr(COLORS, 'background')
        assert hasattr(COLORS, 'surface')
        assert hasattr(COLORS, 'accent')
        assert hasattr(COLORS, 'text_main')
        assert hasattr(COLORS, 'text_dim')

    def test_font_constants(self):
        """Test font constants are defined."""
        assert hasattr(FONTS, 'main')
        assert hasattr(FONTS, 'bold')
        assert hasattr(FONTS, 'header')
        assert hasattr(FONTS, 'sub')
        assert hasattr(FONTS, 'button')

    def test_button_initial_states(self, root):
        """Test buttons have correct initial states."""
        ui = ScheduleAppUI(root)

        # Print button should be enabled
        assert ui.print_btn.cget("state") == "normal"

        # Cancel button should be disabled initially
        assert ui.cancel_btn.cget("state") == "disabled"


@needs_display
@pytest.mark.skipif(not HAS_WIN32PRINT, reason="win32print not available")
class TestUIWithWin32Print:
    """Tests that require win32print."""

    @patch('src.ui.win32print')
    def test_printer_enumeration(self, mock_win32print, root):
        """Test printer enumeration with win32print."""
        # Mock printer enumeration
        mock_win32print.EnumPrinters.side_effect = [
            [("Printer1", "Local", "Printer1")],
            [("Printer2", "Network", "Printer2")]
        ]

        ui = ScheduleAppUI(root)

        # Should have enumerated printers
        assert mock_win32print.EnumPrinters.call_count == 2

    @patch('src.ui.win32print')
    def test_printer_enumeration_error(self, mock_win32print, root):
        """Test printer enumeration error handling."""
        mock_win32print.EnumPrinters.side_effect = Exception("Enumeration error")

        ui = ScheduleAppUI(root)  # Should not crash

        # Printer dropdown should still exist
        assert ui.printer_dropdown is not None


@needs_display
class TestUIMethods:
    """Tests for specific UI methods."""

    def test_browse_folder(self, root):
        """Test folder browse dialog."""
        ui = ScheduleAppUI(root)

        with patch('src.ui.filedialog.askdirectory', return_value="/selected/path"):
            ui._browse_folder(ui.day_entry)

            assert ui.day_entry.get() == "/selected/path"

    def test_browse_folder_cancelled(self, root):
        """Test folder browse dialog when cancelled."""
        ui = ScheduleAppUI(root)
        original_path = "/original/path"
        ui.day_entry.delete(0, tk.END)
        ui.day_entry.insert(0, original_path)

        with patch('src.ui.filedialog.askdirectory', return_value=""):
            ui._browse_folder(ui.day_entry)

            # Path should remain unchanged
            assert ui.day_entry.get() == original_path

    def test_on_printer_change_triggers_callback(self, root):
        """Test printer change triggers config callback."""
        ui = ScheduleAppUI(root)
        mock_callback = MagicMock()
        ui.set_config_change_callback(mock_callback)

        ui._on_printer_change()

        mock_callback.assert_called_once()
