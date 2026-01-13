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
    DATE_PLACEHOLDER,
    COM_ERROR_RPC_CALL_REJECTED,
    COM_ERROR_RPC_SERVERCALL_RETRYLATER,
    PRINTER_STATUS_OFFLINE,
    PRINTER_STATUS_ERROR
)
from .logger import get_logger
from .path_validation import validate_folder_path

logger = get_logger(__name__)



class WordProcessor:
    """Handles Word document operations via COM automation."""

    def __init__(self):
        """Initialize WordProcessor."""
        self.word_app: Optional[Any] = None
        self._initialized = False
        self._thread_id: Optional[int] = None  # Track which thread initialized COM

    def initialize(self) -> None:
        """
        Initialize the Word application instance.

        Raises:
            RuntimeError: If Word cannot be initialized or platform is incompatible
        """
        if self._initialized:
            return

        if not HAS_PYWIN32 or pythoncom is None:
            raise RuntimeError(
                "This application requires Windows with pywin32 installed. "
                "Current platform: " + sys.platform
            )

        com_initialized = False
        try:
            pythoncom.CoInitialize()  # type: ignore
            com_initialized = True
            self._thread_id = threading.get_ident()  # Record which thread initialized COM
            self.word_app = win32com.client.Dispatch("Word.Application")  # type: ignore
            if self.word_app:
                self.word_app.Visible = False
                self.word_app.DisplayAlerts = 0
                # Set macro security to disable macros for security
                # wdSecurityPolicy = 4 (Disable all macros without notification)
                self.word_app.AutomationSecurity = 4
                self._initialized = True
                logger.info(f"Word application initialized on thread {self._thread_id}")
            else:
                raise RuntimeError("win32com.client.Dispatch returned None")
        except Exception as e:
            # Clean up COM if it was initialized but Word creation failed
            if com_initialized and pythoncom:
                try:
                    pythoncom.CoUninitialize()
                except Exception as cleanup_error:
                    logger.warning(f"Error during COM cleanup after failed initialization: {cleanup_error}")
            logger.error(f"Failed to initialize Word application: {e}")
            raise RuntimeError(f"Could not initialize Word: {e}") from e

    def shutdown(self) -> None:
        """Shutdown the Word application instance."""
        if not self.word_app:
            return

        # Warn if shutdown is called from a different thread than initialization
        current_thread = threading.get_ident()
        if self._thread_id is not None and current_thread != self._thread_id:
            logger.warning(
                f"shutdown() called from thread {current_thread}, but COM was initialized on thread {self._thread_id}. "
                f"This may cause COM cleanup issues."
            )

        try:
            self.word_app.Quit()
            logger.info("Word application shut down")
        except Exception as e:
            logger.warning(f"Error shutting down Word: {e}")
        finally:
            self.word_app = None
            self._initialized = False
            if pythoncom:
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
            retries: Number of retry attempts (minimum 1, will always try at least once)
            delay: Delay between retries in seconds

        Returns:
            The result of the function call

        Raises:
            RuntimeError: If called from wrong thread
            Exception: If all retry attempts fail
        """
        # Enforce thread affinity for COM objects
        current_thread = threading.get_ident()
        if self._thread_id is not None and current_thread != self._thread_id:
            raise RuntimeError(
                f"COM call attempted from wrong thread. "
                f"Initialized on thread {self._thread_id}, called from thread {current_thread}. "
                f"COM objects must be used on the same thread they were created on."
            )

        # Ensure at least one attempt
        max_attempts = max(1, retries)
        last_error = None

        for attempt in range(max_attempts):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last_error = e
                error_str = str(e).lower()
                # "Call was rejected by callee" is a common transient COM error
                if ("rejected" in error_str or
                    COM_ERROR_RPC_SERVERCALL_RETRYLATER.lower() in error_str or
                    COM_ERROR_RPC_CALL_REJECTED.lower() in error_str):
                    if attempt < max_attempts - 1:
                        logger.debug(f"COM call rejected, retrying ({attempt + 1}/{max_attempts}) in {delay}s...")
                        time.sleep(delay)
                        continue

                # Non-retriable error, fail immediately
                logger.error(f"COM call failed with non-retriable error: {e}")
                raise e

        # All retry attempts exhausted for rejection errors
        if last_error:
            logger.error(f"COM call failed after {max_attempts} attempts (all rejected)")
            raise last_error

        # This should never happen (loop should always execute at least once)
        raise RuntimeError("COM call failed unexpectedly")

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
        if not self._initialized or not self.word_app:
            return False, "Word processor not initialized"

        # Edge Case: Verify printer is ready/online (Windows only)
        if HAS_PYWIN32 and win32print:
            try:
                phandle = win32print.OpenPrinter(printer_name)
                try:
                    pinfo = win32print.GetPrinter(phandle, 2)
                    status = pinfo.get('Status', 0)
                    if status & PRINTER_STATUS_OFFLINE:
                        return False, f"Printer '{printer_name}' is offline."
                    if status & PRINTER_STATUS_ERROR:
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

            if not doc:
                return False, f"Failed to open document: {target_file}"

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
            # Ensure document is closed (use simple call to avoid masking original error)
            if doc:
                try:
                    # Use direct COM call without retries to fail fast in cleanup
                    doc.Close(CLOSE_NO_SAVE)
                except Exception as e:
                    # Log but don't raise - we don't want cleanup to mask the original error
                    logger.warning(f"Error closing document in cleanup: {e}")

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
        # Note: We use specific patterns with day/month names to reduce false positives
        # Build pattern that matches actual day and month names
        day_names = "(Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday)"
        month_names = "(January|February|March|April|May|June|July|August|September|October|November|December)"

        patterns = [
            # Style 1: "Sunday, January 04, 2026" (with comma after day)
            (
                f"{day_names}, {month_names} [0-9]{{1,2}}, [0-9]{{4}}",
                f"{new_day}, {new_month} {new_day_num}, {new_year}"
            ),
            # Style 2: "Saturday January 03, 2026" (no comma after day)
            (
                f"{day_names} {month_names} [0-9]{{1,2}}, [0-9]{{4}}",
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
        if not success:
            if is_wildcard:
                # Wildcard patterns not matching is expected (template may use different date format)
                logger.debug(f"Wildcard pattern not found: {find_text[:50]}...")
            else:
                # Explicit placeholder not found - worth logging as it may indicate misconfigured template
                logger.info(f"Placeholder not found in document: {find_text}")

    def __enter__(self):
        """Context manager entry."""
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.shutdown()
