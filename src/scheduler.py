"""
Date and scheduling logic for Shift Automator application.

This module handles date calculations, including special scheduling rules
like third Thursday detection.
"""

import calendar
import functools
from datetime import date, timedelta
from typing import Optional

from .constants import THURSDAY
from .logger import get_logger

logger = get_logger(__name__)


@functools.lru_cache(maxsize=None)
def get_third_thursday(year: int, month: int) -> Optional[int]:
    """
    Get the day of the third Thursday in the specified month and year.
    Cached to improve performance.

    Args:
        year: The year
        month: The month (1-12)

    Returns:
        The day of the month that is the third Thursday, or None if not found.
    """
    month_calendar = calendar.monthcalendar(year, month)
    thursdays = [week[THURSDAY] for week in month_calendar if week[THURSDAY] != 0]

    if len(thursdays) >= 3:
        return thursdays[2]
    return None


def is_third_thursday(dt: date) -> bool:
    """
    Check if a given date is the third Thursday of its month.

    Args:
        dt: The date to check

    Returns:
        True if the date is the third Thursday, False otherwise

    Examples:
        >>> is_third_thursday(date(2026, 1, 15))  # Third Thursday of January 2026
        True
        >>> is_third_thursday(date(2026, 1, 8))   # Second Thursday
        False
    """
    if dt.weekday() != THURSDAY:
        return False

    third_thursday = get_third_thursday(dt.year, dt.month)

    # Check if this is the third Thursday
    if third_thursday is not None:
        is_third = dt.day == third_thursday
        logger.debug(f"Date {dt} is third Thursday: {is_third}")
        return is_third

    return False


def get_shift_template_name(dt: date, shift_type: str = "day") -> str:
    """
    Get the template name for a given date and shift type.

    Args:
        dt: The date to get the template for
        shift_type: Either "day" or "night"

    Returns:
        The template name (e.g., "Thursday", "Thursday Night", "THIRD Thursday")

    Raises:
        ValueError: If shift_type is not "day" or "night"
    """
    if shift_type not in ("day", "night"):
        raise ValueError(f"shift_type must be 'day' or 'night', got '{shift_type}'")

    day_name = dt.strftime("%A")

    if shift_type == "day":
        # Day shift uses "THIRD Thursday" for third Thursdays
        if is_third_thursday(dt):
            template_name = "THIRD Thursday"
            logger.info(f"Day shift for {dt}: {template_name} (special schedule)")
        else:
            template_name = day_name
            logger.debug(f"Day shift for {dt}: {template_name}")
    else:
        # Night shift always uses "DayName Night"
        template_name = f"{day_name} Night"
        logger.debug(f"Night shift for {dt}: {template_name}")

    return template_name


def validate_date_range(start_date: date, end_date: date) -> tuple[bool, Optional[str]]:
    """
    Validate that a date range is valid.

    Args:
        start_date: The start date
        end_date: The end date

    Returns:
        Tuple of (is_valid, error_message)
    """
    if end_date < start_date:
        return False, "End date cannot be before start date"
    return True, None


def get_date_range(start_date: date, end_date: date) -> list[date]:
    """
    Get a list of dates in the given range (inclusive).

    Args:
        start_date: The start date
        end_date: The end date

    Returns:
        List of dates from start_date to end_date (inclusive)

    Raises:
        ValueError: If end_date is before start_date
    """
    is_valid, error_msg = validate_date_range(start_date, end_date)
    if not is_valid:
        raise ValueError(error_msg)

    delta = (end_date - start_date).days
    return [start_date + timedelta(days=i) for i in range(delta + 1)]
