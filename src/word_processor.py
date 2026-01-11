"""
Word document processing for Shift Automator application.

This module handles all interactions with Microsoft Word via COM automation,
including document opening, date replacement, and printing.
"""

import os
import time
from datetime import date
from pathlib import Path
from typing import Optional, Any, Tuple

import pythoncom
import win32com.client

from .constants import (
    DOCX_EXTENSION,
    PROTECTION_NONE,
    CLOSE_NO_SAVE,
    COM_RETRIES,
    COM_RETRY_DELAY
)
from .logger import get_logger
from .path_validation import validate_folder_path, is_path_within_base

logger = get_logger(__name__)


class WordProcessor:
    """Handles Word document operations via COM automation."""

    def __init__(self):
        """Initialize WordProcessor."""
        self.word_app: Optional[Any] = None
        self._initialized = False

    def initialize(self) -> None:
        """
        Initialize the Word application instance.

        Raises:
            RuntimeError: If Word cannot be initialized
        """
        if self._initialized:
            return

        try:
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = 0
            self._initialized = True
            logger.info("Word application initialized")
        except Exception as e:
            logger.error(f"Failed to initialize Word application: {e}")
            raise RuntimeError(f"Could not initialize Word: {e}") from e

    def shutdown(self) -> None:
        """Shutdown the Word application instance."""
        if self.word_app:
            try:
                self.word_app.Quit()
                logger.info("Word application shut down")
            except Exception as e:
                logger.warning(f"Error shutting down Word: {e}")
            finally:
                self.word_app = None
                self._initialized = False
                pythoncom.CoUninitialize()

    def safe_com_call(self, func: callable, *args: Any, retries: int = COM_RETRIES,
                      delay: float = COM_RETRY_DELAY) -> Any:
        """
        Execute a COM call with retry logic for transient errors.

        Args:
            func: The COM function to call
            *args: Arguments to pass to the function
            retries: Number of retry attempts
            delay: Delay between retries in seconds

        Returns:
            The result of the function call

        Raises:
            Exception: If all retry attempts fail
        """
        for attempt in range(retries):
            try:
                return func(*args)
            except Exception as e:
                error_str = str(e).lower()
                if "rejected" in error_str or "call was rejected" in error_str:
                    if attempt < retries - 1:
                        logger.debug(f"COM call rejected, retrying ({attempt + 1}/{retries})")
                        time.sleep(delay)
                        continue
                logger.error(f"COM call failed after {attempt + 1} attempts: {e}")
                raise

    def find_template_file(self, folder: str, template_name: str) -> Optional[str]:
        """
        Find a template file in the given folder.

        First tries exact match, then falls back to partial match.

        Args:
            folder: The folder to search in
            template_name: The name of the template (without extension)

        Returns:
            Full path to the template file, or None if not found
        """
        # Validate folder path
        is_valid, error_msg = validate_folder_path(folder)
        if not is_valid:
            logger.error(f"Invalid folder path: {error_msg}")
            return None

        try:
            files = os.listdir(folder)
            exact_match = None
            partial_match = None

            # Search for .docx files
            for f in files:
                if not f.lower().endswith(DOCX_EXTENSION):
                    continue

                # Try exact match first
                if f.lower() == f"{template_name.lower()}{DOCX_EXTENSION}":
                    exact_match = os.path.join(folder, f)
                    logger.debug(f"Found exact match: {exact_match}")
                    break

                # Fall back to partial match
                if template_name.lower() in f.lower() and partial_match is None:
                    partial_match = os.path.join(folder, f)
                    logger.debug(f"Found partial match: {partial_match}")

            target_file = exact_match or partial_match

            if target_file:
                logger.info(f"Found template: {target_file}")
            else:
                logger.warning(f"Template not found: {template_name} in {folder}")

            return target_file

        except OSError as e:
            logger.error(f"Error listing files in {folder}: {e}")
            return None

    def print_document(self, folder: str, template_name: str, current_date: date,
                       printer_name: str) -> Tuple[bool, Optional[str]]:
        """
        Open, update dates, and print a Word document.

        Args:
            folder: The folder containing the template
            template_name: The name of the template file
            current_date: The date to use for replacements
            printer_name: The printer to use

        Returns:
            Tuple of (success, error_message)
        """
        if not self._initialized:
            return False, "Word processor not initialized"

        # Find the template file
        target_file = self.find_template_file(folder, template_name)
        if not target_file:
            return False, f"Template not found: {template_name}"

        doc = None
        try:
            # Open the document
            logger.debug(f"Opening document: {target_file}")
            doc = self.safe_com_call(
                self.word_app.Documents.Open,
                target_file, False, False
            )

            # Unprotect if necessary
            if doc.ProtectionType != PROTECTION_NONE:
                try:
                    self.safe_com_call(doc.Unprotect)
                    logger.debug("Document unprotected")
                except Exception as e:
                    logger.warning(f"Could not unprotect document: {e}")

            # Replace dates
            self.replace_dates(doc, current_date)

            # Set printer and print
            self.word_app.ActivePrinter = printer_name
            logger.debug(f"Printing to: {printer_name}")
            self.safe_com_call(doc.PrintOut, Background=False)

            # Close document
            self.safe_com_call(doc.Close, CLOSE_NO_SAVE)
            doc = None

            logger.info(f"Successfully printed: {template_name}")
            return True, None

        except Exception as e:
            logger.error(f"Error printing document {target_file}: {e}")
            return False, str(e)

        finally:
            # Ensure document is closed
            if doc:
                try:
                    self.safe_com_call(doc.Close, CLOSE_NO_SAVE)
                except Exception as e:
                    logger.warning(f"Error closing document: {e}")

    def replace_dates(self, doc: Any, current_date: date) -> None:
        """
        Replace date placeholders in the document using regex patterns.

        Args:
            doc: The Word document object
            current_date: The date to use for replacements
        """
        # Format date components
        new_day = current_date.strftime("%A")
        new_month = current_date.strftime("%B")
        new_day_num = str(int(current_date.strftime("%d")))
        new_year = current_date.strftime("%Y")

        # Patterns to replace (using Word wildcard syntax)
        # [A-Za-z]@ means "one or more letters"
        patterns = [
            # Day Shift Style (With Comma): "Sunday, January 04, 2026"
            (
                "[A-Za-z]@, [A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_day}, {new_month} {new_day_num}, {new_year}"
            ),
            # Night Shift Style (No Comma): "Saturday January 03, 2026"
            (
                "[A-Za-z]@ [A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_day} {new_month} {new_day_num}, {new_year}"
            ),
            # Fallback/Standard Style: "January 04, 2026"
            (
                "[A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_month} {new_day_num}, {new_year}"
            ),
        ]

        for find_text, replace_text in patterns:
            self._execute_replace(doc, find_text, replace_text)

        logger.debug(f"Date replacements completed for {current_date}")

    def _execute_replace(self, doc: Any, find_text: str, replace_text: str) -> None:
        """
        Execute a find and replace operation across all story ranges.

        Args:
            doc: The Word document object
            find_text: The text pattern to find
            replace_text: The replacement text
        """
        try:
            for story in doc.StoryRanges:
                self._run_find_replace(story, find_text, replace_text)
                nxt = story.NextStoryRange
                while nxt:
                    self._run_find_replace(nxt, find_text, replace_text)
                    nxt = nxt.NextStoryRange
        except Exception as e:
            logger.warning(f"Error during find/replace: {e}")

    def _run_find_replace(self, range_obj: Any, find_text: str, replace_text: str) -> None:
        """
        Run a single find and replace operation on a range.

        Args:
            range_obj: The Word range object
            find_text: The text pattern to find
            replace_text: The replacement text
        """
        try:
            f = range_obj.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            # Execute: FindText, MatchCase, MatchWholeWord, MatchWildcards,
            #          MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format,
            #          ReplaceWith, Replace
            f.Execute(
                find_text,      # FindText
                False,          # MatchCase
                False,          # MatchWholeWord
                True,           # MatchWildcards
                False,          # MatchSoundsLike
                False,          # MatchAllWordForms
                True,           # Forward
                1,              # Wrap (wdFindContinue)
                False,          # Format
                replace_text,   # ReplaceWith
                2               # Replace (wdReplaceAll)
            )
        except Exception as e:
            logger.warning(f"Error in find/replace operation: {e}")

    def __enter__(self):
        """Context manager entry."""
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.shutdown()
