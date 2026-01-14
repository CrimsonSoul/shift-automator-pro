"""
Configuration management for Shift Automator application.

This module handles loading, saving, and validating configuration settings.
"""

import json
import os
import tempfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, Any, Optional

from .constants import CONFIG_FILENAME
from .logger import get_logger
from .utils import get_app_data_dir

logger = get_logger(__name__)


def _get_default_config_dir() -> Path:
    """
    Get the default config directory based on the operating system.

    Returns:
        Path to the default config directory
    """
    return get_app_data_dir("ShiftAutomator")


@dataclass
class AppConfig:
    """Application configuration data class."""
    day_folder: str = ""
    night_folder: str = ""
    printer_name: str = ""

    def to_dict(self) -> Dict[str, str]:
        """Convert config to dictionary."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'AppConfig':
        """Create config from dictionary."""
        return cls(
            day_folder=data.get("day_folder", ""),
            night_folder=data.get("night_folder", ""),
            printer_name=data.get("printer_name", "")
        )

    def validate(self) -> tuple[bool, Optional[str]]:
        """
        Validate configuration values.

        Returns:
            Tuple of (is_valid, error_message)
        """
        if self.day_folder and not Path(self.day_folder).is_dir():
            return False, f"Day folder does not exist: {self.day_folder}"
        if self.night_folder and not Path(self.night_folder).is_dir():
            return False, f"Night folder does not exist: {self.night_folder}"
        return True, None


class ConfigManager:
    """Manages application configuration loading and saving."""

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize ConfigManager.

        Args:
            config_path: Path to config file (default: CONFIG_FILENAME in AppData directory)
        """
        if config_path:
            self.config_path = Path(config_path)
        else:
            config_dir = _get_default_config_dir()
            config_dir.mkdir(parents=True, exist_ok=True)
            self.config_path = config_dir / CONFIG_FILENAME
        self._config: Optional[AppConfig] = None

    def load(self) -> AppConfig:
        """
        Load configuration from file.

        Returns:
            AppConfig instance with loaded values or defaults

        Raises:
            json.JSONDecodeError: If config file contains invalid JSON
            IOError: If config file cannot be read
        """
        if not self.config_path.exists():
            logger.info(f"Config file not found at {self.config_path}, using defaults")
            self._config = AppConfig()
            return self._config

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self._config = AppConfig.from_dict(data)
                logger.info(f"Configuration loaded from {self.config_path}")
                return self._config
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON in config file: {e}")
            raise
        except IOError as e:
            logger.error(f"Error reading config file: {e}")
            raise

    def save(self, config: Optional[AppConfig] = None) -> None:
        """
        Save configuration to file using atomic write pattern.

        Args:
            config: AppConfig instance to save (uses current config if None)

        Raises:
            IOError: If config file cannot be written
        """
        config_to_save = config or self._config
        if config_to_save is None:
            logger.warning("No configuration to save")
            return

        try:
            # Ensure parent directory exists
            self.config_path.parent.mkdir(parents=True, exist_ok=True)

            # Atomic write: write to temp file first, then rename
            fd, tmp_path = tempfile.mkstemp(
                dir=self.config_path.parent,
                suffix='.tmp'
            )
            try:
                with os.fdopen(fd, 'w', encoding='utf-8') as f:
                    json.dump(config_to_save.to_dict(), f, indent=4)
                # os.replace is atomic on POSIX, near-atomic on Windows
                os.replace(tmp_path, self.config_path)
            except Exception:
                # Clean up temp file on failure
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass
                raise

            logger.info(f"Configuration saved to {self.config_path}")
        except IOError as e:
            logger.error(f"Error saving config file: {e}")
            raise

    @property
    def config(self) -> AppConfig:
        """Get current configuration."""
        if self._config is None:
            self._config = self.load()
        return self._config

    @config.setter
    def config(self, value: AppConfig) -> None:
        """Set current configuration."""
        self._config = value
