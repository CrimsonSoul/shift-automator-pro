"""
Integration tests for main application module.
"""

import sys
import threading
from datetime import date
from unittest.mock import MagicMock, patch

import pytest

# Import the class directly, then grab the actual module from sys.modules
# (src.main as a name is shadowed by the main() function exported in src/__init__.py)
from src.main import ShiftAutomatorApp

main_module = sys.modules["src.main"]


class TestShiftAutomatorApp:
    """Tests for ShiftAutomatorApp class."""

    @pytest.fixture
    def app(self):
        """Create a ShiftAutomatorApp with mocked UI and dependencies."""
        with patch.object(main_module, "ScheduleAppUI") as MockUI, \
             patch.object(main_module, "ConfigManager") as MockConfig:
            mock_root = MagicMock()
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = "/tmp/day"
            mock_ui.get_night_folder.return_value = "/tmp/night"
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 14)
            mock_ui.get_end_date.return_value = date(2026, 1, 14)
            mock_ui.progress_var = MagicMock()
            mock_ui.progress_var.get.return_value = 0.0
            mock_ui.print_btn = MagicMock()

            mock_config = MockConfig.return_value
            mock_config.load.return_value = MagicMock(
                day_folder="", night_folder="", printer_name=""
            )

            app = ShiftAutomatorApp(mock_root)
            yield app

    def test_validate_inputs_missing_day_folder(self, app):
        """Should fail validation when day folder is empty."""
        app.ui.get_day_folder.return_value = ""
        is_valid, error = app._validate_inputs()
        assert is_valid is False
        assert "Day Templates" in error

    def test_validate_inputs_missing_night_folder(self, app):
        """Should fail validation when night folder is empty."""
        app.ui.get_night_folder.return_value = ""
        is_valid, error = app._validate_inputs()
        assert is_valid is False
        assert "Night Templates" in error

    def test_validate_inputs_missing_printer(self, app):
        """Should fail validation when no printer selected."""
        app.ui.get_printer_name.return_value = "Choose Printer"
        is_valid, error = app._validate_inputs()
        assert is_valid is False
        assert "printer" in error.lower()

    def test_validate_inputs_missing_dates(self, app):
        """Should fail validation when dates are missing."""
        app.ui.get_start_date.return_value = None
        with patch.object(main_module, "validate_folder_path", return_value=(True, None)):
            is_valid, error = app._validate_inputs()
        assert is_valid is False
        assert "date" in error.lower()

    @patch.object(main_module, "WordProcessor")
    @patch.object(main_module, "validate_folder_path", return_value=(True, None))
    def test_process_batch_success(self, mock_validate, mock_wp_class, app):
        """Should process all days and report success."""
        mock_wp = MagicMock()
        mock_wp.print_document.return_value = (True, None)
        mock_wp.__enter__ = MagicMock(return_value=mock_wp)
        mock_wp.__exit__ = MagicMock(return_value=False)
        mock_wp_class.return_value = mock_wp

        params = {
            'start_date': date(2026, 1, 14),
            'end_date': date(2026, 1, 14),
            'day_folder': '/tmp/day',
            'night_folder': '/tmp/night',
            'printer_name': 'Test Printer',
        }

        app._process_batch(params)

        # Should print both day and night shift
        assert mock_wp.print_document.call_count == 2

    @patch.object(main_module, "WordProcessor")
    @patch.object(main_module, "validate_folder_path", return_value=(True, None))
    def test_process_batch_cancel(self, mock_validate, mock_wp_class, app):
        """Should stop processing when cancel event is set."""
        mock_wp = MagicMock()
        mock_wp.print_document.return_value = (True, None)
        mock_wp.__enter__ = MagicMock(return_value=mock_wp)
        mock_wp.__exit__ = MagicMock(return_value=False)
        mock_wp_class.return_value = mock_wp

        # Set cancel before processing starts
        app._cancel_event.set()

        params = {
            'start_date': date(2026, 1, 14),
            'end_date': date(2026, 1, 16),
            'day_folder': '/tmp/day',
            'night_folder': '/tmp/night',
            'printer_name': 'Test Printer',
        }

        app._process_batch(params)

        # Should not have printed anything (cancelled immediately)
        assert mock_wp.print_document.call_count == 0

    @patch.object(main_module, "WordProcessor")
    @patch.object(main_module, "validate_folder_path", return_value=(True, None))
    def test_process_batch_tracks_failures(self, mock_validate, mock_wp_class, app):
        """Should track failed operations."""
        mock_wp = MagicMock()
        # Day shift fails, night shift succeeds
        mock_wp.print_document.side_effect = [
            (False, "Template not found"),
            (True, None),
        ]
        mock_wp.__enter__ = MagicMock(return_value=mock_wp)
        mock_wp.__exit__ = MagicMock(return_value=False)
        mock_wp_class.return_value = mock_wp

        params = {
            'start_date': date(2026, 1, 14),
            'end_date': date(2026, 1, 14),
            'day_folder': '/tmp/day',
            'night_folder': '/tmp/night',
            'printer_name': 'Test Printer',
        }

        app._process_batch(params)

        # Verify failure summary was scheduled to show
        app.root.after.assert_called()

    def test_on_close_without_active_thread(self, app):
        """Should destroy window immediately if no thread is running."""
        app._processing_thread = None
        app._on_close()
        app.root.destroy.assert_called_once()

    def test_on_close_with_active_thread(self, app):
        """Should cancel and join thread before destroying window."""
        mock_thread = MagicMock()
        mock_thread.is_alive.return_value = True
        app._processing_thread = mock_thread

        app._on_close()

        assert app._cancel_event.is_set()
        mock_thread.join.assert_called_once_with(timeout=5)
        app.root.destroy.assert_called_once()
