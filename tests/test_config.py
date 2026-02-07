"""
Unit tests for config module.
"""

import json
import os
import tempfile
from pathlib import Path

import pytest

from src.config import AppConfig, ConfigManager


class TestAppConfig:
    """Tests for AppConfig dataclass."""

    def test_default_values(self):
        """Default config should have empty strings."""
        config = AppConfig()
        assert config.day_folder == ""
        assert config.night_folder == ""
        assert config.printer_name == ""
        assert config.headers_footers_only is False

    def test_with_values(self):
        """Config should store provided values."""
        config = AppConfig(
            day_folder="/path/to/day",
            night_folder="/path/to/night",
            printer_name="My Printer",
            headers_footers_only=True,
        )
        assert config.day_folder == "/path/to/day"
        assert config.night_folder == "/path/to/night"
        assert config.printer_name == "My Printer"
        assert config.headers_footers_only is True

    def test_to_dict(self):
        """Config should convert to dictionary."""
        config = AppConfig(
            day_folder="/path/to/day",
            night_folder="/path/to/night",
            printer_name="My Printer",
            headers_footers_only=True,
        )
        result = config.to_dict()
        assert result == {
            "day_folder": "/path/to/day",
            "night_folder": "/path/to/night",
            "printer_name": "My Printer",
            "headers_footers_only": True,
        }

    def test_from_dict(self):
        """Config should create from dictionary."""
        data = {
            "day_folder": "/path/to/day",
            "night_folder": "/path/to/night",
            "printer_name": "My Printer",
            "headers_footers_only": True,
        }
        config = AppConfig.from_dict(data)
        assert config.day_folder == "/path/to/day"
        assert config.night_folder == "/path/to/night"
        assert config.printer_name == "My Printer"
        assert config.headers_footers_only is True

    def test_from_dict_with_missing_keys(self):
        """Config should use defaults for missing keys."""
        data = {"day_folder": "/path/to/day"}
        config = AppConfig.from_dict(data)
        assert config.day_folder == "/path/to/day"
        assert config.night_folder == ""
        assert config.printer_name == ""
        assert config.headers_footers_only is False

    def test_validate_with_valid_folders(self, tmp_path):
        """Validation should pass with valid folders."""
        day_folder = tmp_path / "day"
        night_folder = tmp_path / "night"
        day_folder.mkdir()
        night_folder.mkdir()

        config = AppConfig(
            day_folder=str(day_folder),
            night_folder=str(night_folder),
            printer_name="My Printer",
        )
        is_valid, error = config.validate()
        assert is_valid is True
        assert error is None

    def test_validate_with_invalid_day_folder(self):
        """Validation should fail with non-existent day folder."""
        config = AppConfig(
            day_folder="/nonexistent/path", night_folder="", printer_name="My Printer"
        )
        is_valid, error = config.validate()
        assert is_valid is False
        assert error is not None
        assert "does not exist" in error

    def test_validate_with_invalid_night_folder(self):
        """Validation should fail with non-existent night folder."""
        config = AppConfig(
            day_folder="", night_folder="/nonexistent/path", printer_name="My Printer"
        )
        is_valid, error = config.validate()
        assert is_valid is False
        assert error is not None
        assert "does not exist" in error

    def test_validate_with_empty_folders(self):
        """Validation should pass with empty folders."""
        config = AppConfig(day_folder="", night_folder="", printer_name="My Printer")
        is_valid, error = config.validate()
        assert is_valid is True
        assert error is None


class TestConfigManager:
    """Tests for ConfigManager class."""

    def test_load_nonexistent_file(self, tmp_path):
        """Loading non-existent file should return default config."""
        config_file = tmp_path / "config.json"
        manager = ConfigManager(str(config_file))
        config = manager.load()
        assert isinstance(config, AppConfig)
        assert config.day_folder == ""
        assert config.night_folder == ""
        assert config.printer_name == ""

    def test_save_and_load(self, tmp_path):
        """Saving and loading should preserve config."""
        config_file = tmp_path / "config.json"
        manager = ConfigManager(str(config_file))

        original_config = AppConfig(
            day_folder="/path/to/day",
            night_folder="/path/to/night",
            printer_name="My Printer",
        )
        manager.save(original_config)

        loaded_config = manager.load()
        assert loaded_config.day_folder == original_config.day_folder
        assert loaded_config.night_folder == original_config.night_folder
        assert loaded_config.printer_name == original_config.printer_name

    def test_load_invalid_json(self, tmp_path):
        """Loading invalid JSON should raise JSONDecodeError."""
        config_file = tmp_path / "config.json"
        config_file.write_text("{ invalid json }")

        manager = ConfigManager(str(config_file))
        with pytest.raises(json.JSONDecodeError):
            manager.load()

    def test_config_property(self, tmp_path):
        """Config property should load and cache config."""
        config_file = tmp_path / "config.json"
        manager = ConfigManager(str(config_file))

        # First access should load
        config1 = manager.config
        assert isinstance(config1, AppConfig)

        # Second access should return cached
        config2 = manager.config
        assert config1 is config2

    def test_config_setter(self, tmp_path):
        """Config setter should update cached config."""
        config_file = tmp_path / "config.json"
        manager = ConfigManager(str(config_file))

        new_config = AppConfig(
            day_folder="/new/path",
            night_folder="/new/night",
            printer_name="New Printer",
        )
        manager.config = new_config

        assert manager.config.day_folder == "/new/path"
        assert manager.config is new_config
