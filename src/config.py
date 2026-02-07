"""
Configuration management for Shift Automator application.

This module handles loading, saving, and validating configuration settings.
"""

import json
import os
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, Any, Optional

from .constants import CONFIG_FILENAME
from .app_paths import get_data_dir
from .logger import get_logger

__all__ = ["AppConfig", "ConfigManager"]

logger = get_logger(__name__)


@dataclass
class AppConfig:
    """Application configuration data class."""

    day_folder: str = ""
    night_folder: str = ""
    printer_name: str = ""
    headers_footers_only: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Convert config to dictionary."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AppConfig":
        """Create config from dictionary."""
        return cls(
            day_folder=data.get("day_folder", ""),
            night_folder=data.get("night_folder", ""),
            printer_name=data.get("printer_name", ""),
            headers_footers_only=bool(data.get("headers_footers_only", False)),
        )

    def validate(self) -> tuple[bool, Optional[str]]:
        """
        Validate configuration values.

        Returns:
            Tuple of (is_valid, error_message)
        """
        if self.day_folder and not os.path.isdir(self.day_folder):
            return False, f"Day folder does not exist: {self.day_folder}"
        if self.night_folder and not os.path.isdir(self.night_folder):
            return False, f"Night folder does not exist: {self.night_folder}"
        return True, None


class ConfigManager:
    """Manages application configuration loading and saving."""

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize ConfigManager.

        Args:
            config_path: Path to config file (default: CONFIG_FILENAME in current directory)
        """
        self._legacy_config_path = Path(CONFIG_FILENAME)
        self._allow_legacy_migration = config_path is None
        if config_path:
            self.config_path = Path(config_path)
        else:
            self.config_path = get_data_dir() / CONFIG_FILENAME
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
            # Backward-compatibility: older versions stored config.json in the working directory.
            if self._allow_legacy_migration and self._legacy_config_path.exists():
                try:
                    with open(self._legacy_config_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        self._config = AppConfig.from_dict(data)
                        logger.info(
                            f"Configuration loaded from legacy path {self._legacy_config_path}; "
                            f"migrating to {self.config_path}"
                        )
                    try:
                        self.save(self._config)
                    except Exception as e:
                        logger.warning(
                            f"Could not migrate legacy config to {self.config_path}: {e}"
                        )
                    # Rename old file so it doesn't get picked up on the next launch.
                    try:
                        migrated = self._legacy_config_path.with_suffix(
                            ".json.migrated"
                        )
                        self._legacy_config_path.rename(migrated)
                        logger.info(f"Legacy config renamed to {migrated}")
                    except Exception as e:
                        logger.debug(f"Could not rename legacy config: {e}")
                    return self._config
                except Exception as e:
                    logger.warning(
                        f"Could not load legacy config at {self._legacy_config_path}: {e}"
                    )

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
        Save configuration to file.

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

            # Write atomically: write to a temp file then replace.
            tmp_path = self.config_path.with_suffix(self.config_path.suffix + ".tmp")
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(config_to_save.to_dict(), f, indent=4)
                f.flush()
                os.fsync(f.fileno())

            os.replace(tmp_path, self.config_path)
            logger.info(f"Configuration saved to {self.config_path}")
        except (IOError, OSError) as e:
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
