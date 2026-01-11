"""
Unit tests for constants module.

These tests verify that all constants are properly defined and have expected values.
"""

import pytest

from src.constants import (
    MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY, SUNDAY,
    PROTECTION_NONE, PROTECTION_READ_ONLY, PROTECTION_ALLOW_COMMENTS,
    PROTECTION_ALLOW_REVISIONS,
    CLOSE_NO_SAVE, CLOSE_SAVE, CLOSE_PROMPT,
    PRINTER_ENUM_LOCAL, PRINTER_ENUM_NETWORK,
    DOCX_EXTENSION,
    CONFIG_FILENAME, LOG_FILENAME,
    WINDOW_WIDTH, WINDOW_HEIGHT, WINDOW_RESIZABLE,
    PROGRESS_MAX,
    COM_RETRIES, COM_RETRY_DELAY, COM_TIMEOUT,
    COLORS, FONTS
)


class TestWeekdayConstants:
    """Test weekday constant values."""

    def test_monday_value(self):
        """Test Monday constant value."""
        assert MONDAY == 0

    def test_tuesday_value(self):
        """Test Tuesday constant value."""
        assert TUESDAY == 1

    def test_wednesday_value(self):
        """Test Wednesday constant value."""
        assert WEDNESDAY == 2

    def test_thursday_value(self):
        """Test Thursday constant value."""
        assert THURSDAY == 3

    def test_friday_value(self):
        """Test Friday constant value."""
        assert FRIDAY == 4

    def test_saturday_value(self):
        """Test Saturday constant value."""
        assert SATURDAY == 5

    def test_sunday_value(self):
        """Test Sunday constant value."""
        assert SUNDAY == 6

    def test_weekday_sequential(self):
        """Test weekday constants are sequential."""
        weekdays = [MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY, SUNDAY]
        assert weekdays == list(range(7))


class TestWordProtectionConstants:
    """Test Word document protection constants."""

    def test_protection_none(self):
        """Test PROTECTION_NONE value."""
        assert PROTECTION_NONE == -1

    def test_protection_read_only(self):
        """Test PROTECTION_READ_ONLY value."""
        assert PROTECTION_READ_ONLY == 0

    def test_protection_allow_comments(self):
        """Test PROTECTION_ALLOW_COMMENTS value."""
        assert PROTECTION_ALLOW_COMMENTS == 1

    def test_protection_allow_revisions(self):
        """Test PROTECTION_ALLOW_REVISIONS value."""
        assert PROTECTION_ALLOW_REVISIONS == 2


class TestWordCloseConstants:
    """Test Word document close options."""

    def test_close_no_save(self):
        """Test CLOSE_NO_SAVE value."""
        assert CLOSE_NO_SAVE == 0

    def test_close_save(self):
        """Test CLOSE_SAVE value."""
        assert CLOSE_SAVE == 1

    def test_close_prompt(self):
        """Test CLOSE_PROMPT value."""
        assert CLOSE_PROMPT == 2


class TestPrinterConstants:
    """Test printer enumeration constants."""

    def test_printer_enum_local(self):
        """Test PRINTER_ENUM_LOCAL value."""
        assert PRINTER_ENUM_LOCAL == 2

    def test_printer_enum_network(self):
        """Test PRINTER_ENUM_NETWORK value."""
        assert PRINTER_ENUM_NETWORK == 4


class TestFileConstants:
    """Test file-related constants."""

    def test_docx_extension(self):
        """Test DOCX extension."""
        assert DOCX_EXTENSION == ".docx"
        assert DOCX_EXTENSION.startswith(".")

    def test_config_filename(self):
        """Test config filename."""
        assert CONFIG_FILENAME == "config.json"
        assert CONFIG_FILENAME.endswith(".json")

    def test_log_filename(self):
        """Test log filename."""
        assert LOG_FILENAME == "shift_automator.log"
        assert LOG_FILENAME.endswith(".log")


class TestUIConstants:
    """Test UI-related constants."""

    def test_window_width(self):
        """Test window width."""
        assert isinstance(WINDOW_WIDTH, int)
        assert WINDOW_WIDTH == 640

    def test_window_height(self):
        """Test window height."""
        assert isinstance(WINDOW_HEIGHT, int)
        assert WINDOW_HEIGHT == 720

    def test_window_resizable(self):
        """Test window resizable setting."""
        assert isinstance(WINDOW_RESIZABLE, bool)
        assert WINDOW_RESIZABLE is False


class TestProgressConstants:
    """Test progress-related constants."""

    def test_progress_max(self):
        """Test maximum progress value."""
        assert PROGRESS_MAX == 100


class TestCOMConstants:
    """Test COM operation constants."""

    def test_com_retries(self):
        """Test COM retry count."""
        assert isinstance(COM_RETRIES, int)
        assert COM_RETRIES == 5

    def test_com_retry_delay(self):
        """Test COM retry delay."""
        assert isinstance(COM_RETRY_DELAY, (int, float))
        assert COM_RETRY_DELAY == 1

    def test_com_timeout(self):
        """Test COM timeout value."""
        assert isinstance(COM_TIMEOUT, (int, float))
        assert COM_TIMEOUT == 30


class TestColorConstants:
    """Test color scheme constants."""

    def test_colors_has_background(self):
        """Test background color exists."""
        assert hasattr(COLORS, 'background')
        assert COLORS.background.startswith('#')
        assert len(COLORS.background) == 7

    def test_colors_has_surface(self):
        """Test surface color exists."""
        assert hasattr(COLORS, 'surface')
        assert COLORS.surface.startswith('#')

    def test_colors_has_accent(self):
        """Test accent color exists."""
        assert hasattr(COLORS, 'accent')
        assert COLORS.accent.startswith('#')

    def test_colors_has_text_main(self):
        """Test main text color exists."""
        assert hasattr(COLORS, 'text_main')
        assert COLORS.text_main.startswith('#')

    def test_colors_has_text_dim(self):
        """Test dim text color exists."""
        assert hasattr(COLORS, 'text_dim')
        assert COLORS.text_dim.startswith('#')

    def test_colors_has_border(self):
        """Test border color exists."""
        assert hasattr(COLORS, 'border')
        assert COLORS.border.startswith('#')

    def test_colors_are_dataclass(self):
        """Test COLORS is a dataclass instance."""
        assert hasattr(COLORS, '__dataclass_fields__')

    def test_all_required_colors_present(self):
        """Test all expected color attributes are present."""
        expected_colors = [
            'background', 'surface', 'accent', 'text_main', 'text_dim',
            'success', 'border', 'secondary', 'accent_hover'
        ]
        for color in expected_colors:
            assert hasattr(COLORS, color), f"Missing color: {color}"


class TestFontConstants:
    """Test font configuration constants."""

    def test_fonts_has_main(self):
        """Test main font exists."""
        assert hasattr(FONTS, 'main')
        assert isinstance(FONTS.main, tuple)
        assert len(FONTS.main) == 2

    def test_fonts_has_bold(self):
        """Test bold font exists."""
        assert hasattr(FONTS, 'bold')
        assert isinstance(FONTS.bold, tuple)
        assert len(FONTS.bold) == 3

    def test_fonts_has_header(self):
        """Test header font exists."""
        assert hasattr(FONTS, 'header')
        assert isinstance(FONTS.header, tuple)
        assert len(FONTS.header) == 3

    def test_fonts_has_sub(self):
        """Test sub font exists."""
        assert hasattr(FONTS, 'sub')
        assert isinstance(FONTS.sub, tuple)
        assert len(FONTS.sub) == 2

    def test_fonts_has_button(self):
        """Test button font exists."""
        assert hasattr(FONTS, 'button')
        assert isinstance(FONTS.button, tuple)
        assert len(FONTS.button) == 3

    def test_fonts_are_dataclass(self):
        """Test FONTS is a dataclass instance."""
        assert hasattr(FONTS, '__dataclass_fields__')

    def test_font_names_are_strings(self):
        """Test font names are strings."""
        assert isinstance(FONTS.main[0], str)
        assert isinstance(FONTS.bold[0], str)
        assert isinstance(FONTS.header[0], str)
        assert isinstance(FONTS.sub[0], str)
        assert isinstance(FONTS.button[0], str)

    def test_font_sizes_are_numbers(self):
        """Test font sizes are numbers."""
        assert isinstance(FONTS.main[1], int)
        assert isinstance(FONTS.bold[1], int)
        assert isinstance(FONTS.header[1], int)
        assert isinstance(FONTS.sub[1], int)
        assert isinstance(FONTS.button[1], int)

    def test_font_families_match(self):
        """Test all fonts use the same font family."""
        font_family = FONTS.main[0]
        assert FONTS.bold[0] == font_family
        assert FONTS.header[0] == font_family
        assert FONTS.sub[0] == font_family
        assert FONTS.button[0] == font_family
