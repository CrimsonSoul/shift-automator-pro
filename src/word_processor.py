"""
Word document processing for Shift Automator application.

This module handles all interactions with Microsoft Word via COM automation,
including document opening, date replacement, and printing.
"""

import os
import time
import re
from datetime import date
from pathlib import Path
from typing import Optional, Any, Tuple, Callable

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
        self.word_app = None
        self._initialized = False
        self._template_cache: dict[str, dict[str, str]] = {}

    def initialize(self) -> None:
        """
        Initialize the Word application instance.

        Raises:
            RuntimeError: If Word cannot be initialized
        """
        if self._initialized and self.word_app:
            return

        try:
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            if self.word_app:
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
                try:
                    pythoncom.CoUninitialize()
                except Exception as e:
                    logger.debug(f"Error in CoUninitialize: {e}")

    def clear_template_cache(self, folder: Optional[str] = None) -> None:
        """
        Clear the template cache.

        Args:
            folder: Specific folder to clear, or None to clear all
        """
        if folder:
            folder_path = str(Path(folder).resolve())
            self._template_cache.pop(folder_path, None)
            logger.debug(f"Cleared template cache for: {folder_path}")
        else:
            self._template_cache.clear()
            logger.debug("Cleared all template caches")

    def safe_com_call(self, func: Callable[..., Any], *args: Any, retries: int = COM_RETRIES,
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

        Uses caching for faster lookup and robust matching logic.

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

        folder_path = str(Path(folder).resolve())
        template_name_lower = " ".join(template_name.lower().split())

        # Build cache if not already present for this folder
        if folder_path not in self._template_cache:
            try:
                files = os.listdir(folder_path)
                cache = {}
                for f in files:
                    if f.lower().endswith(DOCX_EXTENSION):
                        base_name = " ".join(f.lower().replace(DOCX_EXTENSION, "").split())
                        cache[base_name] = os.path.join(folder_path, f)
                self._template_cache[folder_path] = cache
                logger.debug(f"Cached {len(cache)} templates from {folder_path}")
            except OSError as e:
                logger.error(f"Error listing files in {folder_path}: {e}")
                return None

        cache = self._template_cache[folder_path]
        logger.debug(f"Looking for template '{template_name}' (normalized: '{template_name_lower}') "
                     f"in cache with {len(cache)} entries: {list(cache.keys())}")

        # 1. Try exact match
        if template_name_lower in cache:
            target = cache[template_name_lower]
            logger.debug(f"Found exact template match: {target}")
            return target

        # 2. Try robust matching using word boundaries
        # This prevents "Thursday" matching "THIRD Thursday"
        # but allows "Thursday" matching "Thursday Night" if it's the only match
        logger.debug(f"No exact match for '{template_name_lower}', trying robust matching")
        pattern = re.compile(rf"\b{re.escape(template_name_lower)}\b")
        
        matches = []
        for base_name, full_path in cache.items():
            if pattern.search(base_name):
                # Special logic: if search term doesn't have "third" but filename does, skip
                # This prevents "Thursday" matching "THIRD Thursday"
                if "third" not in template_name_lower and "third" in base_name:
                    continue
                matches.append(full_path)

        if len(matches) == 1:
            logger.info(f"Found robust template match: {matches[0]}")
            return matches[0]
        elif len(matches) > 1:
            # If multiple matches, try to find the one that starts with it (more specific)
            for m in matches:
                if Path(m).stem.lower().startswith(template_name_lower):
                    logger.info(f"Found specific template match from multiple: {m}")
                    return m
            
            logger.warning(f"Ambiguous template matches for '{template_name}': {matches}")
            return matches[0] # Fallback to first

        logger.warning(f"Template not found: {template_name} in {folder}")
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

        # Find the template file
        target_file = self.find_template_file(folder, template_name)
        if not target_file:
            return False, f"Template not found: {template_name}"

        # Verify template is within the expected folder (prevents path traversal)
        if not is_path_within_base(target_file, folder):
            logger.error(f"Template path '{target_file}' is outside folder '{folder}'")
            return False, f"Template path is outside the expected folder"
        logger.info(f"Template '{template_name}' resolved to: {target_file}")

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
            if self.word_app:
                self.word_app.ActivePrinter = printer_name
            logger.debug(f"Printing to: {printer_name}")
            # PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, ...)
            # Background=False ensures synchronous printing
            self.safe_com_call(doc.PrintOut, False)


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
        # Normalize non-breaking spaces before running patterns
        self._normalize_spaces_in_doc(doc)

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
            # Day Shift Style Abbreviated: "Sun, January 04, 2026"
            (
                "[A-Za-z]{3}, [A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
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

        any_matched = False
        for find_text, replace_text in patterns:
            if self._execute_replace(doc, find_text, replace_text):
                any_matched = True

        if not any_matched:
            logger.warning(f"No date patterns matched in document for {current_date}. "
                          f"The template may use an unsupported date format.")

        logger.debug(f"Date replacements completed for {current_date}")

    def _normalize_spaces_in_doc(self, doc: Any) -> None:
        """
        Replace non-breaking spaces with regular spaces in the document.

        Word templates frequently contain non-breaking spaces (Unicode 00A0)
        which prevent wildcard find/replace patterns from matching date strings.
        """
        try:
            for story in doc.StoryRanges:
                f = story.Find
                f.ClearFormatting()
                f.Replacement.ClearFormatting()
                f.Execute(
                    "^s",       # FindText: ^s = non-breaking space in Word
                    False,      # MatchCase
                    False,      # MatchWholeWord
                    False,      # MatchWildcards (must be False for special codes)
                    False,      # MatchSoundsLike
                    False,      # MatchAllWordForms
                    True,       # Forward
                    1,          # Wrap (wdFindContinue)
                    False,      # Format
                    " ",        # ReplaceWith: regular space
                    2           # Replace (wdReplaceAll)
                )
        except Exception as e:
            logger.debug(f"Non-breaking space normalization: {e}")

    def _execute_replace(self, doc: Any, find_text: str, replace_text: str) -> bool:
        """
        Execute a find and replace operation across all story ranges.

        Args:
            doc: The Word document object
            find_text: The text pattern to find
            replace_text: The replacement text

        Returns:
            True if at least one replacement was made
        """
        any_replaced = False
        try:
            for story in doc.StoryRanges:
                if self._run_find_replace(story, find_text, replace_text):
                    any_replaced = True
                nxt = story.NextStoryRange
                while nxt:
                    if self._run_find_replace(nxt, find_text, replace_text):
                        any_replaced = True
                    nxt = nxt.NextStoryRange
        except Exception as e:
            logger.warning(f"Error during find/replace: {e}")
        return any_replaced

    def _run_find_replace(self, range_obj: Any, find_text: str, replace_text: str) -> bool:
        """
        Run a single find and replace operation on a range.

        Args:
            range_obj: The Word range object
            find_text: The text pattern to find
            replace_text: The replacement text

        Returns:
            True if the pattern was found and replaced
        """
        try:
            f = range_obj.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            # Execute: FindText, MatchCase, MatchWholeWord, MatchWildcards,
            #          MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format,
            #          ReplaceWith, Replace
            result = f.Execute(
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
            if result:
                logger.debug(f"Find/replace matched: '{find_text}' -> '{replace_text}'")
            return bool(result)
        except Exception as e:
            logger.warning(f"Error in find/replace operation: {e}")
            return False

    def __enter__(self):
        """Context manager entry."""
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.shutdown()
