"""
Integration tests for Shift Automator application.

These tests verify the end-to-end functionality of the main application workflow.
"""

import os
import tempfile
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch, Mock
import pytest
import tkinter as tk

from src.main import ShiftAutomatorApp
from src.config import AppConfig


class TestShiftAutomatorAppIntegration:
    """Integration tests for ShiftAutomatorApp."""

    @pytest.fixture
    def temp_folders(self, tmp_path):
        """Create temporary folders for day and night templates."""
        day_folder = tmp_path / "day_templates"
        night_folder = tmp_path / "night_templates"
        day_folder.mkdir()
        night_folder.mkdir()

        # Create template files
        template_days = ["Monday", "Tuesday", "Wednesday", "Thursday",
                        "THIRD Thursday", "Friday", "Saturday", "Sunday"]

        for day in template_days:
            (day_folder / f"{day}.docx").write_text(f"Day template for {day}")
            if day != "THIRD Thursday":  # Night templates don't have "THIRD Thursday"
                (night_folder / f"{day} Night.docx").write_text(f"Night template for {day}")

        return {"day": str(day_folder), "night": str(night_folder)}

    @pytest.fixture
    def mock_root(self):
        """Create a mock Tkinter root window."""
        root = MagicMock(spec=tk.Tk)
        root.after = MagicMock()
        return root

    def test_app_initialization(self, mock_root, temp_folders):
        """Test that the application initializes properly."""
        with patch('src.main.ScheduleAppUI'):
            app = ShiftAutomatorApp(mock_root)

            assert app.root == mock_root
            assert app.ui is not None
            assert app.config_manager is not None
            assert app._cancel_requested is False

    def test_config_persistence(self, mock_root, temp_folders, tmp_path):
        """Test that configuration is saved and loaded correctly."""
        config_file = tmp_path / "test_config.json"

        with patch('src.main.ScheduleAppUI'), \
             patch('src.config.ConfigManager.config_path', config_file):

            app = ShiftAutomatorApp(mock_root)

            # Save configuration
            test_config = AppConfig(
                day_folder=temp_folders["day"],
                night_folder=temp_folders["night"],
                printer_name="Test Printer"
            )
            app._save_config(test_config)

            # Verify config file was created
            assert config_file.exists()

            # Load configuration
            loaded_config = app.config_manager.load()
            assert loaded_config.day_folder == temp_folders["day"]
            assert loaded_config.night_folder == temp_folders["night"]
            assert loaded_config.printer_name == "Test Printer"

    def test_input_validation(self, mock_root, temp_folders):
        """Test that input validation works correctly."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 1)
            mock_ui.get_end_date.return_value = date(2026, 1, 5)

            app = ShiftAutomatorApp(mock_root)

            # Test valid inputs
            is_valid, error_msg = app._validate_inputs()
            assert is_valid is True
            assert error_msg is None

    def test_input_validation_missing_folders(self, mock_root):
        """Test validation fails when folders are missing."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = ""
            mock_ui.get_night_folder.return_value = ""
            mock_ui.get_printer_name.return_value = "Test Printer"

            app = ShiftAutomatorApp(mock_root)

            is_valid, error_msg = app._validate_inputs()
            assert is_valid is False
            assert "Day Templates folder" in error_msg

    def test_input_validation_invalid_date_range(self, mock_root, temp_folders):
        """Test validation fails when end date is before start date."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 10)
            mock_ui.get_end_date.return_value = date(2026, 1, 1)

            app = ShiftAutomatorApp(mock_root)

            is_valid, error_msg = app._validate_inputs()
            assert is_valid is False
            assert "before start date" in error_msg

    def test_batch_processing_cancellation(self, mock_root, temp_folders):
        """Test that batch processing can be cancelled."""
        with patch('src.main.ScheduleAppUI') as MockUI, \
             patch('src.main.WordProcessor') as MockWordProc:

            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 1)
            mock_ui.get_end_date.return_value = date(2026, 1, 10)

            # Mock Word processor
            mock_word_instance = MockWordProc.return_value.__enter__.return_value
            mock_word_instance.print_document.return_value = (True, None)

            app = ShiftAutomatorApp(mock_root)

            # Start processing
            app.start_processing()

            # Request cancellation immediately
            app.cancel_processing()

            assert app._cancel_requested is True
            mock_ui.set_cancel_button_state.assert_called_with("disabled")

    def test_start_processing_updates_ui_state(self, mock_root, temp_folders):
        """Test that starting processing updates UI button states."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 1)
            mock_ui.get_end_date.return_value = date(2026, 1, 1)

            app = ShiftAutomatorApp(mock_root)
            app.start_processing()

            # Verify UI state changes
            mock_ui.set_print_button_state.assert_called_with("disabled")
            mock_ui.set_cancel_button_state.assert_called_with("normal")
            assert app._cancel_requested is False

    @patch('src.main.WordProcessor')
    def test_batch_processing_single_day(self, MockWordProc, mock_root, temp_folders):
        """Test processing a single day."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 6)  # Monday
            mock_ui.get_end_date.return_value = date(2026, 1, 6)

            # Mock Word processor
            mock_word_instance = MockWordProc.return_value.__enter__.return_value
            mock_word_instance.print_document.return_value = (True, None)

            app = ShiftAutomatorApp(mock_root)

            # Process in main thread for testing
            app._process_batch()

            # Verify print_document was called twice (day and night shift)
            assert mock_word_instance.print_document.call_count == 2

            # Verify correct templates were requested
            calls = mock_word_instance.print_document.call_args_list
            assert calls[0][0][1] == "Monday"  # Day shift
            assert calls[1][0][1] == "Monday Night"  # Night shift

    @patch('src.main.WordProcessor')
    def test_batch_processing_third_thursday(self, MockWordProc, mock_root, temp_folders):
        """Test processing third Thursday uses special template."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 15)  # Third Thursday
            mock_ui.get_end_date.return_value = date(2026, 1, 15)

            # Mock Word processor
            mock_word_instance = MockWordProc.return_value.__enter__.return_value
            mock_word_instance.print_document.return_value = (True, None)

            app = ShiftAutomatorApp(mock_root)
            app._process_batch()

            # Verify correct templates were used
            calls = mock_word_instance.print_document.call_args_list
            assert calls[0][0][1] == "THIRD Thursday"  # Day shift special template
            assert calls[1][0][1] == "Thursday Night"  # Night shift regular

    @patch('src.main.WordProcessor')
    def test_batch_processing_failure_tracking(self, MockWordProc, mock_root, temp_folders):
        """Test that failures are tracked and reported."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 6)
            mock_ui.get_end_date.return_value = date(2026, 1, 6)

            # Mock Word processor to fail
            mock_word_instance = MockWordProc.return_value.__enter__.return_value
            mock_word_instance.print_document.return_value = (False, "Template not found")

            app = ShiftAutomatorApp(mock_root)
            app._process_batch()

            # Verify show_warning was called for failures (this gets called on after())
            # Since we're in the main thread, after() executes immediately
            assert mock_root.after.called

    def test_config_change_callback(self, mock_root, temp_folders):
        """Test that configuration changes are saved via callback."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"

            app = ShiftAutomatorApp(mock_root)

            # Trigger config change callback
            app._save_current_config()

            # Verify config was updated
            config = app.config_manager.load()
            assert config.day_folder == temp_folders["day"]
            assert config.night_folder == temp_folders["night"]

    @patch('src.main.WordProcessor')
    def test_batch_processing_multiple_days(self, MockWordProc, mock_root, temp_folders):
        """Test processing multiple days in sequence."""
        with patch('src.main.ScheduleAppUI') as MockUI:
            mock_ui = MockUI.return_value
            mock_ui.get_day_folder.return_value = temp_folders["day"]
            mock_ui.get_night_folder.return_value = temp_folders["night"]
            mock_ui.get_printer_name.return_value = "Test Printer"
            mock_ui.get_start_date.return_value = date(2026, 1, 6)  # Monday
            mock_ui.get_end_date.return_value = date(2026, 1, 8)    # Wednesday (3 days)

            # Mock Word processor
            mock_word_instance = MockWordProc.return_value.__enter__.return_value
            mock_word_instance.print_document.return_value = (True, None)

            app = ShiftAutomatorApp(mock_root)
            app._process_batch()

            # Verify print_document was called 6 times (3 days × 2 shifts)
            assert mock_word_instance.print_document.call_count == 6

            # Verify correct sequence of templates
            calls = mock_word_instance.print_document.call_args_list
            expected_templates = [
                "Monday", "Monday Night",
                "Tuesday", "Tuesday Night",
                "Wednesday", "Wednesday Night"
            ]
            actual_templates = [call[0][1] for call in calls]
            assert actual_templates == expected_templates
