"""
Word document processing for Shift Automator application.

This module handles all interactions with Microsoft Word via COM automation,
including document opening, date replacement, and printing.
"""

import os
import sys
import threading
import time
from datetime import date
from pathlib import Path
from typing import Optional, Any, Tuple, Callable

# Platform-specific imports
try:
    import pythoncom
    import win32com.client
    import win32print
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False
    pythoncom = None  # type: ignore
    win32com = None  # type: ignore
    win32print = None # type: ignore

from .constants import (
    DOCX_EXTENSION,
    PROTECTION_NONE,
    CLOSE_NO_SAVE,
    COM_RETRIES,
    COM_RETRY_DELAY,
    DATE_PLACEHOLDER
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
            RuntimeError: If Word cannot be initialized or platform is incompatible
        """
        if self._initialized:
            return

        if not HAS_PYWIN32:
            raise RuntimeError(
                "This application requires Windows with pywin32 installed. "
                "Current platform: " + sys.platform
            )

        try:
            pythoncom.CoInitialize()  # type: ignore
            self.word_app = win32com.client.Dispatch("Word.Application")  # type: ignore
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = 0
            # Set macro security to disable macros for security
            # wdSecurityPolicy = 4 (Disable all macros without notification)
            self.word_app.AutomationSecurity = 4
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

    def safe_com_call(self, func: Callable[..., Any], *args: Any,
                      retries: int = COM_RETRIES, delay: float = COM_RETRY_DELAY,
                      **kwargs: Any) -> Any:
        """
        Execute a COM call with retry logic for transient errors (like "Call rejected").
        
        Note: This MUST be called from the same thread that initialized Word.
        
        Args:
            func: The COM function to call
            *args: Arguments to pass to the function
            **kwargs: Keyword arguments to pass to the function
            retries: Number of retry attempts
            delay: Delay between retries in seconds

        Returns:
            The result of the function call

        Raises:
            Exception: If all retry attempts fail
        """
        last_error = None
        for attempt in range(retries):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last_error = e
                error_str = str(e).lower()
                # "Call was rejected by callee" is a common transient COM error
                if "rejected" in error_str or "0x80010101" in error_str or "0x80010001" in error_str:
                    if attempt < retries - 1:
                        logger.debug(f"COM call rejected, retrying ({attempt + 1}/{retries}) in {delay}s...")
                        time.sleep(delay)
                        continue
                
                logger.error(f"COM call failed: {e}")
                raise e

        if last_error:
            raise last_error
        raise RuntimeError("COM call failed after all retries")

    def find_template_file(self, folder: str, template_name: str) -> Optional[str]:
        """
        Find a template file in the given folder using exact name matching.

        Searches for files with exact case-insensitive match of template name.

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

            # Construct expected filename
            expected_filename = f"{template_name}{DOCX_EXTENSION}"

            # Search for exact match (case-insensitive)
            for f in files:
                if not f.lower().endswith(DOCX_EXTENSION):
                    continue

                # Exact match (case-insensitive)
                if f.lower() == expected_filename.lower():
                    target_file = os.path.join(folder, f)
                    
                    # Edge Case: Check for empty files
                    try:
                        if os.path.getsize(target_file) == 0:
                            logger.error(f"Template file is empty: {target_file}")
                            return None
                    except OSError as e:
                        logger.error(f"Could not check file size for {target_file}: {e}")
                        return None
                        
                    logger.info(f"Found template: {target_file}")
                    return target_file

            # No match found
            logger.warning(f"Template not found: {expected_filename} in {folder}")
            return None

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

        # Edge Case: Verify printer is ready/online (Windows only)
        if HAS_PYWIN32 and win32print:
            try:
                # PRINTER_STATUS_OFFLINE = 0x00000080
                # PRINTER_STATUS_ERROR = 0x00000002
                phandle = win32print.OpenPrinter(printer_name)
                try:
                    pinfo = win32print.GetPrinter(phandle, 2)
                    status = pinfo.get('Status', 0)
                    if status & 0x00000080: # Offline
                        return False, f"Printer '{printer_name}' is offline."
                    if status & 0x00000002: # Error
                        return False, f"Printer '{printer_name}' reported an error state."
                finally:
                    win32print.ClosePrinter(phandle)
            except Exception as e:
                logger.warning(f"Could not verify printer status for '{printer_name}': {e}")
                # We continue anyway as some drivers don't report status correctly

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
        Replace date placeholders in the document.
        
        Strategy:
        1. Prioritize placeholder replacement (e.g., {{DATE}}).
        2. Fallback to broad regex only if explicitly allowed (implemented here for backward compatibility).

        Args:
            doc: The Word document object
            current_date: The date to use for replacements
        """
        # Format date components
        new_day = current_date.strftime("%A")
        new_month = current_date.strftime("%B")
        new_day_num = str(int(current_date.strftime("%d")))
        new_year = current_date.strftime("%Y")

        # 1. Primary: Placeholder replacement (SAFE)
        # Format: "Thursday, January 15, 2026"
        placeholder_text = f"{new_day}, {new_month} {new_day_num}, {new_year}"
        self._execute_replace(doc, DATE_PLACEHOLDER, placeholder_text, is_wildcard=False)

        # 2. Fallback: Pattern matching (RISKIER - matches existing dates)
        # Note: We use specific patterns to reduce false positives
        patterns = [
            # Style 1: "Sunday, January 04, 2026"
            (
                "[A-Za-z]@, [A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_day}, {new_month} {new_day_num}, {new_year}"
            ),
            # Style 2: "Saturday January 03, 2026"
            (
                "[A-Za-z]@ [A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_day} {new_month} {new_day_num}, {new_year}"
            ),
        ]

        for find_text, replace_text in patterns:
            self._execute_replace(doc, find_text, replace_text, is_wildcard=True)

        logger.debug(f"Date replacements completed for {current_date}")

    def _execute_replace(self, doc: Any, find_text: str, replace_text: str, is_wildcard: bool = True) -> None:
        """
        Execute a find and replace operation across all story ranges.

        Args:
            doc: The Word document object
            find_text: The text pattern to find
            replace_text: The replacement text
            is_wildcard: Whether to use wildcard matching
        """
        for story in doc.StoryRanges:
            self._run_find_replace(story, find_text, replace_text, is_wildcard)
            nxt = story.NextStoryRange
            while nxt:
                self._run_find_replace(nxt, find_text, replace_text, is_wildcard)
                nxt = nxt.NextStoryRange

    def _run_find_replace(self, range_obj: Any, find_text: str, replace_text: str, is_wildcard: bool = True) -> None:
        """
        Run a single find and replace operation on a range.

        Args:
            range_obj: The Word range object
            find_text: The text pattern to find
            replace_text: The replacement text
            is_wildcard: Whether to use wildcard matching
        """
        f = range_obj.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        # Execute: FindText, MatchCase, MatchWholeWord, MatchWildcards,
        #          MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format,
        #          ReplaceWith, Replace
        success = self.safe_com_call(
            f.Execute,
            find_text,      # FindText
            False,          # MatchCase
            False,          # MatchWholeWord
            is_wildcard,    # MatchWildcards
            False,          # MatchSoundsLike
            False,          # MatchAllWordForms
            True,           # Forward
            1,              # Wrap (wdFindContinue)
            False,          # Format
            replace_text,   # ReplaceWith
            2               # Replace (wdReplaceAll)
        )
        if not success and is_wildcard:
            # Note: Execute returns True if any match was found.
            # We don't necessarily want to raise an error if a wildcard match wasn't found,
            # but for explicit {{DATE}} placeholders, we might want to know.
            pass

    def __enter__(self):
        """Context manager entry."""
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.shutdown()
