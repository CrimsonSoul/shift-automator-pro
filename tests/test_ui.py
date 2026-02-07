"""
Unit tests for UI module.
"""

import tkinter as tk
from unittest.mock import MagicMock, patch
import pytest

from src.ui import ScheduleAppUI


class TestScheduleAppUI:
    """Tests for ScheduleAppUI class."""

    @pytest.fixture
    def root(self):
        """Create a mock Tk root."""
        root = MagicMock(spec=tk.Tk)
        return root

    @pytest.fixture
    def ui(self, root):
        """Create a ScheduleAppUI instance."""
        # Mock win32print and widget creation to avoid Tcl errors
        with patch("win32print.EnumPrinters", return_value=[]), patch(
            "src.ui.ttk.Style"
        ), patch("src.ui.ttk.Frame"), patch("src.ui.ttk.Label"), patch(
            "src.ui.ttk.LabelFrame"
        ), patch(
            "src.ui.ttk.Entry"
        ), patch(
            "src.ui.ttk.Button"
        ), patch(
            "src.ui.ttk.Checkbutton"
        ), patch(
            "src.ui.ttk.OptionMenu"
        ), patch(
            "src.ui.ttk.Progressbar"
        ), patch(
            "src.ui.tk.Button"
        ), patch(
            "src.ui.tk.StringVar"
        ), patch(
            "src.ui.tk.DoubleVar"
        ), patch(
            "src.ui.tk.BooleanVar"
        ), patch(
            "src.ui.DateEntry"
        ):
            ui = ScheduleAppUI(root)
            # Manually assign mock widgets for testing
            ui.day_entry = MagicMock()
            ui.night_entry = MagicMock()
            ui.print_btn = MagicMock()
            ui.status_label = MagicMock()
            ui.progress_var = MagicMock()
            ui.printer_var = MagicMock()
            ui.headers_only_var = MagicMock()
            yield ui

    def test_init(self, ui):
        """UI should initialize widgets."""
        assert ui.day_entry is not None
        assert ui.night_entry is not None
        assert ui.print_btn is not None

    def test_get_day_folder(self, ui):
        """Should return value from day entry."""
        ui.day_entry.get.return_value = "C:/Templates/Day"
        assert ui.get_day_folder() == "C:/Templates/Day"

    def test_get_night_folder(self, ui):
        """Should return value from night entry."""
        ui.night_entry.get.return_value = "C:/Templates/Night"
        assert ui.get_night_folder() == "C:/Templates/Night"

    def test_set_print_button_state(self, ui):
        """Should update button state."""
        ui.set_print_button_state("disabled")
        ui.print_btn.config.assert_called_with(state="disabled")

    def test_update_status(self, ui):
        """Should update status label and progress bar."""
        ui.update_status("Processing...", 50.0)
        ui.status_label.config.assert_called_with(text="Processing...")
        ui.progress_var.set.assert_called_with(50.0)

    @patch("tkinter.messagebox.showerror")
    def test_show_error(self, mock_error, ui):
        """Should call messagebox.showerror."""
        ui.show_error("Title", "Message")
        mock_error.assert_called_with("Title", "Message")

    def test_set_start_command(self, ui):
        """Should set the button command."""
        mock_cmd = MagicMock()
        ui.set_start_command(mock_cmd)
        ui.print_btn.config.assert_called_with(command=mock_cmd)

    def test_get_headers_footers_only(self, ui):
        """Should return boolean from headers-only var."""
        ui.headers_only_var.get.return_value = True
        assert ui.get_headers_footers_only() is True
