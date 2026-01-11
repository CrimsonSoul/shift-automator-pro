"""
Unit tests for scheduler module.
"""

import pytest
from datetime import date

from src.scheduler import is_third_thursday, get_shift_template_name, validate_date_range, get_date_range


class TestIsThirdThursday:
    """Tests for is_third_thursday function."""

    def test_third_thursday_january_2026(self):
        """January 15, 2026 is the third Thursday."""
        assert is_third_thursday(date(2026, 1, 15)) is True

    def test_second_thursday_january_2026(self):
        """January 8, 2026 is the second Thursday."""
        assert is_third_thursday(date(2026, 1, 8)) is False

    def test_fourth_thursday_january_2026(self):
        """January 22, 2026 is the fourth Thursday."""
        assert is_third_thursday(date(2026, 1, 22)) is False

    def test_not_thursday(self):
        """January 15, 2026 is a Thursday, but January 14 is Wednesday."""
        assert is_third_thursday(date(2026, 1, 14)) is False

    def test_third_thursday_february_2026(self):
        """February 19, 2026 is the third Thursday."""
        assert is_third_thursday(date(2026, 2, 19)) is True

    def test_third_thursday_march_2026(self):
        """March 19, 2026 is the third Thursday."""
        assert is_third_thursday(date(2026, 3, 19)) is True


class TestGetShiftTemplateName:
    """Tests for get_shift_template_name function."""

    def test_day_shift_regular_day(self):
        """Regular day shift should use day name."""
        result = get_shift_template_name(date(2026, 1, 14), "day")
        assert result == "Wednesday"

    def test_day_shift_third_thursday(self):
        """Third Thursday day shift should use 'THIRD Thursday'."""
        result = get_shift_template_name(date(2026, 1, 15), "day")
        assert result == "THIRD Thursday"

    def test_night_shift_regular_day(self):
        """Night shift should use 'DayName Night' format."""
        result = get_shift_template_name(date(2026, 1, 14), "night")
        assert result == "Wednesday Night"

    def test_night_shift_third_thursday(self):
        """Night shift on third Thursday should still use 'DayName Night'."""
        result = get_shift_template_name(date(2026, 1, 15), "night")
        assert result == "Thursday Night"

    def test_invalid_shift_type(self):
        """Invalid shift type should raise ValueError."""
        with pytest.raises(ValueError, match="shift_type must be 'day' or 'night'"):
            get_shift_template_name(date(2026, 1, 14), "invalid")


class TestValidateDateRange:
    """Tests for validate_date_range function."""

    def test_valid_range(self):
        """Valid date range should return True."""
        start = date(2026, 1, 1)
        end = date(2026, 1, 31)
        is_valid, error = validate_date_range(start, end)
        assert is_valid is True
        assert error is None

    def test_same_day(self):
        """Same start and end date should be valid."""
        start = date(2026, 1, 15)
        end = date(2026, 1, 15)
        is_valid, error = validate_date_range(start, end)
        assert is_valid is True
        assert error is None

    def test_end_before_start(self):
        """End date before start date should be invalid."""
        start = date(2026, 1, 31)
        end = date(2026, 1, 1)
        is_valid, error = validate_date_range(start, end)
        assert is_valid is False
        assert error == "End date cannot be before start date"


class TestGetDateRange:
    """Tests for get_date_range function."""

    def test_single_day(self):
        """Single day range should return list with one date."""
        start = date(2026, 1, 15)
        end = date(2026, 1, 15)
        result = get_date_range(start, end)
        assert result == [date(2026, 1, 15)]

    def test_multiple_days(self):
        """Multiple day range should return all dates."""
        start = date(2026, 1, 1)
        end = date(2026, 1, 3)
        result = get_date_range(start, end)
        assert result == [
            date(2026, 1, 1),
            date(2026, 1, 2),
            date(2026, 1, 3)
        ]

    def test_week_range(self):
        """Week range should return 7 dates."""
        start = date(2026, 1, 4)  # Monday
        end = date(2026, 1, 10)  # Sunday
        result = get_date_range(start, end)
        assert len(result) == 7
        assert result[0] == date(2026, 1, 4)
        assert result[-1] == date(2026, 1, 10)

    def test_invalid_range_raises_error(self):
        """Invalid range should raise ValueError."""
        start = date(2026, 1, 31)
        end = date(2026, 1, 1)
        with pytest.raises(ValueError, match="End date cannot be before start date"):
            get_date_range(start, end)
