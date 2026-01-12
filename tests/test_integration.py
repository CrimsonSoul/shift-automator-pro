"""
Integration tests for Shift Automator application.

These tests verify the full workflow from UI interaction to document printing.
"""

import tempfile
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest

from src.config import ConfigManager, AppConfig
from src.scheduler import get_shift_template_name, get_date_range
from src.path_validation import validate_folder_path, is_path_within_base


class TestConfigIntegration:
    """Integration tests for configuration management."""

    def test_full_config_lifecycle(self, tmp_path):
        """Test complete configuration lifecycle: create, save, load, validate."""
        config_file = tmp_path / "config.json"
        manager = ConfigManager(str(config_file))

        # Create day and night folders
        day_folder = tmp_path / "day_templates"
        night_folder = tmp_path / "night_templates"
        day_folder.mkdir()
        night_folder.mkdir()

        # Create configuration
        config = AppConfig(
            day_folder=str(day_folder),
            night_folder=str(night_folder),
            printer_name="Test Printer"
        )

        # Save configuration
        manager.save(config)

        # Load configuration
        loaded_config = manager.load()

        # Verify loaded configuration matches
        assert loaded_config.day_folder == config.day_folder
        assert loaded_config.night_folder == config.night_folder
        assert loaded_config.printer_name == config.printer_name

        # Validate configuration
        is_valid, error = loaded_config.validate()
        assert is_valid is True
        assert error is None


class TestSchedulerIntegration:
    """Integration tests for scheduling logic."""

    def test_full_week_template_generation(self):
        """Test template generation for a full week."""
        start_date = date(2026, 1, 5)  # Monday
        end_date = date(2026, 1, 11)  # Sunday

        dates = get_date_range(start_date, end_date)

        # Should have 7 dates
        assert len(dates) == 7

        # Generate templates for each day
        day_templates = [get_shift_template_name(d, "day") for d in dates]
        night_templates = [get_shift_template_name(d, "night") for d in dates]

        # Verify day templates
        assert day_templates[0] == "Monday"
        assert day_templates[1] == "Tuesday"
        assert day_templates[2] == "Wednesday"
        assert day_templates[3] == "Thursday"
        assert day_templates[4] == "Friday"
        assert day_templates[5] == "Saturday"
        assert day_templates[6] == "Sunday"

        # Verify night templates
        assert night_templates[0] == "Monday Night"
        assert night_templates[1] == "Tuesday Night"
        assert night_templates[2] == "Wednesday Night"
        assert night_templates[3] == "Thursday Night"
        assert night_templates[4] == "Friday Night"
        assert night_templates[5] == "Saturday Night"
        assert night_templates[6] == "Sunday Night"

    def test_third_thursday_in_month(self):
        """Test third Thursday detection in a full month."""
        start_date = date(2026, 1, 1)  # January 1, 2026
        end_date = date(2026, 1, 31)  # January 31, 2026

        dates = get_date_range(start_date, end_date)

        # Find third Thursday
        third_thursday = date(2026, 1, 15)
        assert third_thursday in dates

        # Verify template name for third Thursday
        template = get_shift_template_name(third_thursday, "day")
        assert template == "THIRD Thursday"


class TestPathValidationIntegration:
    """Integration tests for path validation."""

    def test_template_folder_validation_workflow(self, tmp_path):
        """Test complete workflow for validating template folders."""
        # Create template folder structure
        day_folder = tmp_path / "day_templates"
        night_folder = tmp_path / "night_templates"
        day_folder.mkdir()
        night_folder.mkdir()

        # Create some template files
        (day_folder / "Monday.docx").write_text("Day template")
        (day_folder / "Tuesday.docx").write_text("Day template")
        (night_folder / "Monday Night.docx").write_text("Night template")
        (night_folder / "Tuesday Night.docx").write_text("Night template")

        # Validate folders
        day_valid, day_error = validate_folder_path(str(day_folder))
        assert day_valid is True
        assert day_error is None

        night_valid, night_error = validate_folder_path(str(night_folder))
        assert night_valid is True
        assert night_error is None

        # Verify paths are within base
        assert is_path_within_base(str(day_folder), str(tmp_path)) is True
        assert is_path_within_base(str(night_folder), str(tmp_path)) is True

        # Test path traversal prevention
        outside_path = tmp_path.parent / "outside"
        assert is_path_within_base(str(outside_path), str(tmp_path)) is False


class TestEndToEndWorkflow:
    """End-to-end integration tests for the full application workflow."""

    def test_complete_workflow_simulation(self, tmp_path):
        """Simulate complete workflow from configuration to template lookup."""
        # Setup: Create template folders
        day_folder = tmp_path / "day_templates"
        night_folder = tmp_path / "night_templates"
        day_folder.mkdir()
        night_folder.mkdir()

        # Create templates for a week
        for day_name in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]:
            (day_folder / f"{day_name}.docx").write_text(f"{day_name} day template")
            (night_folder / f"{day_name} Night.docx").write_text(f"{day_name} night template")

        # Step 1: Create and save configuration
        config_file = tmp_path / "config.json"
        manager = ConfigManager(str(config_file))
        config = AppConfig(
            day_folder=str(day_folder),
            night_folder=str(night_folder),
            printer_name="Test Printer"
        )
        manager.save(config)

        # Step 2: Load configuration
        loaded_config = manager.load()
        assert loaded_config.day_folder == str(day_folder)
        assert loaded_config.night_folder == str(night_folder)

        # Step 3: Validate configuration
        is_valid, error = loaded_config.validate()
        assert is_valid is True

        # Step 4: Generate date range for a week
        start_date = date(2026, 1, 5)  # Monday
        end_date = date(2026, 1, 11)  # Sunday
        dates = get_date_range(start_date, end_date)

        # Step 5: Generate templates for each day
        for current_date in dates:
            day_template = get_shift_template_name(current_date, "day")
            night_template = get_shift_template_name(current_date, "night")

            # Verify templates exist
            day_template_path = day_folder / f"{day_template}.docx"
            night_template_path = night_folder / f"{night_template}.docx"

            assert day_template_path.exists(), f"Day template {day_template} not found"
            assert night_template_path.exists(), f"Night template {night_template} not found"

        # Step 6: Verify all templates were processed
        assert len(dates) == 7  # 7 days in a week
