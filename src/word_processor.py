"""
Word document processing for Shift Automator application.

This module handles all interactions with Microsoft Word via COM automation,
including document opening, date replacement, and printing.
"""

import gc
import time
import re
from datetime import date
from pathlib import Path
from types import TracebackType
from typing import Iterator, Optional, Any, Callable, cast

try:
    import pythoncom as _pythoncom  # type: ignore
    import win32com.client as _win32_client  # type: ignore
except Exception:  # pragma: no cover - validated at runtime on Windows
    _pythoncom = None
    _win32_client = None

pythoncom = cast(Any, _pythoncom)
win32_client = cast(Any, _win32_client)

from .constants import (
    DOCX_EXTENSION,
    PROTECTION_NONE,
    CLOSE_NO_SAVE,
    COM_RETRIES,
    COM_RETRY_DELAY,
    WD_PRIMARY_HEADER_STORY,
    WD_EVEN_PAGES_HEADER_STORY,
    WD_PRIMARY_FOOTER_STORY,
    WD_EVEN_PAGES_FOOTER_STORY,
    WD_FIRST_PAGE_HEADER_STORY,
    WD_FIRST_PAGE_FOOTER_STORY,
    WD_FIND_CONTINUE,
    WD_REPLACE_ALL,
)
from .logger import get_logger
from .path_validation import validate_folder_path, is_path_within_base


logger = get_logger(__name__)


def get_word_automation_status() -> tuple[bool, str]:
    """Return whether Word COM automation dependencies are available.

    This does not guarantee Microsoft Word is installed, but it ensures the
    pywin32 COM bindings are importable.

    Returns:
        Tuple of ``(available, message)``.  When *available* is ``False``,
        *message* describes what is missing.
    """

    if _pythoncom is None or _win32_client is None:
        return (
            False,
            "Microsoft Word automation dependencies are missing. "
            "Install pywin32 and run on Windows with Microsoft Word available.",
        )
    return True, ""


class TemplateLookupError(Exception):
    """Raised when templates cannot be resolved safely (e.g., ambiguity)."""


class WordProcessor:
    """Handles Word document operations via COM automation."""

    def __init__(self) -> None:
        """Initialize WordProcessor.

        The Word COM connection is *not* opened here; call :meth:`initialize`
        (or use the context manager) to start the Word process.
        """
        self.word_app: Any = None
        self._initialized = False
        self._com_initialized = False
        self._template_cache: dict[str, dict[str, str]] = {}

    def initialize(self) -> None:
        """
        Initialize the Word application instance.

        Raises:
            RuntimeError: If Word cannot be initialized
        """
        if self._initialized and self.word_app:
            return

        if _pythoncom is None or _win32_client is None:
            raise RuntimeError(
                "Microsoft Word automation dependencies are missing. "
                "This app requires Windows with pywin32 installed and Microsoft Word available."
            )

        try:
            pythoncom.CoInitialize()
            self._com_initialized = True

            # Prefer DispatchEx to avoid attaching to an existing interactive Word instance.
            dispatch_ex = getattr(win32_client, "DispatchEx", None)
            use_dispatch_ex = callable(dispatch_ex) and getattr(
                dispatch_ex, "__module__", ""
            ).startswith("win32com")
            if use_dispatch_ex:
                dispatch_ex_fn = cast(Callable[[str], Any], dispatch_ex)
                self.word_app = dispatch_ex_fn("Word.Application")
            else:
                self.word_app = win32_client.Dispatch("Word.Application")

            if self.word_app:
                self.word_app.Visible = False
                self.word_app.DisplayAlerts = 0

                # Best-effort hardening: disable macro execution for automated opens.
                # msoAutomationSecurityForceDisable = 3
                try:
                    self.word_app.AutomationSecurity = 3
                except Exception as e:
                    logger.debug(f"Could not set Word AutomationSecurity: {e}")
            self._initialized = True
            logger.info("Word application initialized")
        except Exception as e:
            logger.error(f"Failed to initialize Word application: {e}")
            # If COM was initialized in this thread, uninitialize to avoid leaking.
            if self._com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception as uninit_e:
                    logger.debug(
                        f"Error in CoUninitialize after init failure: {uninit_e}"
                    )
                finally:
                    self._com_initialized = False
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
                # Force-release the COM reference to prevent zombie Word.exe
                # processes when other references (e.g. exception tracebacks) linger.
                self.word_app = None
                self._initialized = False
                gc.collect()

        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.debug(f"Error in CoUninitialize: {e}")
            finally:
                self._com_initialized = False

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

    def _build_template_cache(self, folder_path: str) -> dict[str, str]:
        """Build a normalized template cache for a folder.

        Scans *folder_path* for ``.docx`` files, filtering out Word lock
        files (``~$*``) and hidden files (``.*``).

        Args:
            folder_path: Absolute path to the template folder.

        Returns:
            Dict mapping normalized (lower-cased, whitespace-collapsed) base
            names to their full file paths.
        """

        cache: dict[str, str] = {}
        for entry in Path(folder_path).iterdir():
            name = entry.name
            # Skip Word temp lock files and hidden files
            if name.startswith("~$") or name.startswith("."):
                continue
            if name.lower().endswith(DOCX_EXTENSION):
                base_name = " ".join(name.lower().replace(DOCX_EXTENSION, "").split())
                cache[base_name] = str(entry)
        return cache

    def _ensure_template_cache(
        self, folder_path: str, force_refresh: bool = False
    ) -> None:
        """Ensure the template cache exists; optionally rebuild it.

        Args:
            folder_path: Absolute path to the template folder.
            force_refresh: If True, rebuild even if a cache already exists.

        Raises:
            TemplateLookupError: If the folder cannot be listed.
        """

        if (not force_refresh) and folder_path in self._template_cache:
            return

        try:
            cache = self._build_template_cache(folder_path)
            self._template_cache[folder_path] = cache
            logger.debug(f"Cached {len(cache)} templates from {folder_path}")
        except OSError as e:
            raise TemplateLookupError(
                f"Error listing files in {folder_path}: {e}"
            ) from e

    def safe_com_call(
        self,
        func: Callable[..., Any],
        *args: Any,
        retries: int = COM_RETRIES,
        delay: float = COM_RETRY_DELAY,
    ) -> Any:
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
        if retries < 1:
            raise ValueError("retries must be >= 1")

        for attempt in range(retries):
            try:
                return func(*args)
            except Exception as e:
                error_str = str(e).lower()
                transient_keywords = ("rejected", "call was rejected", "busy", "server")
                if any(kw in error_str for kw in transient_keywords):
                    if attempt < retries - 1:
                        logger.debug(
                            f"COM call rejected, retrying ({attempt + 1}/{retries})"
                        )
                        time.sleep(delay)
                        continue
                logger.error(f"COM call failed after {attempt + 1} attempts: {e}")
                raise

        # Defensive: this path is only reachable if retries == 0, which is disallowed above.
        raise RuntimeError("COM call failed")

    def find_template_file(self, folder: str, template_name: str) -> Optional[str]:
        """
        Find a template file in the given folder.

        Uses caching for faster lookup and robust matching logic.

        Args:
            folder: The folder to search in.
            template_name: The name of the template (without extension).

        Returns:
            Full path to the template file, or ``None`` if not found.

        Raises:
            TemplateLookupError: If the folder is invalid or multiple
                templates match ambiguously.
        """
        # Validate folder path
        is_valid, error_msg = validate_folder_path(folder)
        if not is_valid:
            raise TemplateLookupError(error_msg or "Invalid template folder")

        folder_path = str(Path(folder).resolve())
        template_name_lower = " ".join(template_name.lower().split())

        # Ensure cache exists (and refresh once on miss to pick up newly added templates)
        had_cache = folder_path in self._template_cache
        self._ensure_template_cache(folder_path)

        for attempt in range(2):
            cache = self._template_cache[folder_path]

            # 1. Try exact match
            if template_name_lower in cache:
                target = cache[template_name_lower]
                logger.debug(f"Template exact match: '{template_name}' -> {target}")
                return target

            # 2. Try robust matching using word boundaries
            # This prevents "Thursday" matching "THIRD Thursday"
            # but allows "Thursday" matching "Thursday Night" if it's the only match
            pattern = re.compile(rf"\b{re.escape(template_name_lower)}\b")

            matches: list[str] = []
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
                exact = [
                    m
                    for m in matches
                    if " ".join(Path(m).stem.lower().split()) == template_name_lower
                ]
                if len(exact) == 1:
                    logger.info(
                        f"Found exact-stem template match from multiple: {exact[0]}"
                    )
                    return exact[0]

                starts = [
                    m
                    for m in matches
                    if Path(m).stem.lower().startswith(template_name_lower)
                ]
                if len(starts) == 1:
                    logger.info(
                        f"Found specific template match from multiple: {starts[0]}"
                    )
                    return starts[0]

                raise TemplateLookupError(
                    f"Ambiguous template matches for '{template_name}'. "
                    f"Please rename templates to be unique. Matches: {matches}"
                )

            # Not found: refresh once in case templates were added during runtime.
            # Only refresh if we had a pre-existing cache; if we just built the cache,
            # refreshing again is unlikely to help and only adds I/O.
            if attempt == 0 and had_cache:
                logger.debug(
                    f"Template not found; refreshing cache for {folder_path} and retrying"
                )
                self._ensure_template_cache(folder_path, force_refresh=True)
                continue

            logger.warning(f"Template not found: {template_name} in {folder}")
            return None

        # Defensive: loop always returns, but keep mypy satisfied.
        return None

    def print_document(
        self,
        folder: str,
        template_name: str,
        current_date: date,
        printer_name: str,
        headers_footers_only: bool = False,
    ) -> tuple[bool, Optional[str]]:
        """
        Open, update dates, and print a Word document.

        Args:
            folder: The folder containing the template
            template_name: The name of the template file
            current_date: The date to use for replacements
            printer_name: The printer to use
            headers_footers_only: If True, only replace dates in headers/footers

        Returns:
            tuple of (success, error_message)
        """
        if not self._initialized or not self.word_app:
            return False, "Word processor not initialized"

        # Find the template file
        try:
            target_file = self.find_template_file(folder, template_name)
        except TemplateLookupError as e:
            logger.error(
                f"Template lookup error for '{template_name}' in '{folder}': {e}"
            )
            return False, str(e)
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
                self.word_app.Documents.Open, target_file, False, False
            )

            # Unprotect if necessary
            if doc.ProtectionType != PROTECTION_NONE:
                try:
                    self.safe_com_call(doc.Unprotect)
                    logger.debug("Document unprotected")
                except Exception as e:
                    logger.warning(f"Could not unprotect document: {e}")

            # Replace dates
            self.replace_dates(
                doc, current_date, headers_footers_only=headers_footers_only
            )

            # Set printer and print
            if self.word_app:
                try:
                    self.word_app.ActivePrinter = printer_name
                except Exception as e:
                    logger.warning(
                        f"Could not set ActivePrinter to '{printer_name}': {e}"
                    )
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

    def replace_dates(
        self, doc: Any, current_date: date, headers_footers_only: bool = False
    ) -> None:
        """
        Replace date placeholders in the document using regex patterns.

        Args:
            doc: The Word document object
            current_date: The date to use for replacements
            headers_footers_only: If True, restrict replacements to
                header/footer story ranges only.
        """
        allowed_story_types: Optional[set[int]] = None
        if headers_footers_only:
            allowed_story_types = {
                WD_PRIMARY_HEADER_STORY,
                WD_EVEN_PAGES_HEADER_STORY,
                WD_FIRST_PAGE_HEADER_STORY,
                WD_PRIMARY_FOOTER_STORY,
                WD_EVEN_PAGES_FOOTER_STORY,
                WD_FIRST_PAGE_FOOTER_STORY,
            }

        # Normalize non-breaking spaces before running patterns
        self._normalize_spaces_in_doc(doc, allowed_story_types=allowed_story_types)

        # Format date components
        new_day = current_date.strftime("%A")
        new_month = current_date.strftime("%B")
        new_day_num = str(int(current_date.strftime("%d")))
        new_year = current_date.strftime("%Y")

        # Patterns to replace (using Word wildcard syntax)
        # [A-Za-z]@ means "one or more letters"
        #
        # IMPORTANT: These pattern groups are ordered from most specific to
        # least specific.  Within each group, once a pattern matches we skip
        # the remaining patterns in that group to prevent overlapping
        # replacements (e.g. the "no comma" pattern re-matching text that
        # was already replaced by the "with comma" pattern, which caused
        # day-name duplication like "SaturSaturday").

        # --- Group 1: Full date with day name (mutually exclusive) ---
        # Try comma-separated first, then abbreviated, then no-comma.
        day_name_patterns = [
            # Day Shift Style (With Comma): "Sunday, January 04, 2026"
            (
                "[A-Za-z]@, [A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_day}, {new_month} {new_day_num}, {new_year}",
            ),
            # Day Shift Style Abbreviated: "Sun, January 04, 2026"
            (
                "[A-Za-z]{3}, [A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_day}, {new_month} {new_day_num}, {new_year}",
            ),
            # Night Shift Style (No Comma): "Saturday January 03, 2026"
            (
                "[A-Za-z]@ [A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_day} {new_month} {new_day_num}, {new_year}",
            ),
        ]

        # --- Group 2: Date without day name (fallback) ---
        fallback_patterns = [
            # Fallback/Standard Style: "January 04, 2026"
            (
                "[A-Za-z]@ [0-9]{1,2}, [0-9]{4}",
                f"{new_month} {new_day_num}, {new_year}",
            ),
        ]

        any_matched = False

        # Run day-name patterns: stop at the first one that matches to
        # avoid later, broader patterns re-matching already-replaced text.
        for find_text, replace_text in day_name_patterns:
            if self._execute_replace(
                doc, find_text, replace_text, allowed_story_types=allowed_story_types
            ):
                any_matched = True
                break

        # Run fallback patterns only if no day-name pattern matched, since
        # the fallback pattern is a substring of the day-name patterns and
        # would corrupt text that was already correctly replaced above.
        if not any_matched:
            for find_text, replace_text in fallback_patterns:
                if self._execute_replace(
                    doc,
                    find_text,
                    replace_text,
                    allowed_story_types=allowed_story_types,
                ):
                    any_matched = True

        if not any_matched:
            logger.warning(
                f"No date patterns matched in document for {current_date}. "
                f"The template may use an unsupported date format."
            )

        logger.debug(f"Date replacements completed for {current_date}")

    def _normalize_spaces_in_doc(
        self, doc: Any, allowed_story_types: Optional[set[int]] = None
    ) -> None:
        """Replace non-breaking spaces with regular spaces in the document.

        Word templates frequently contain non-breaking spaces (Unicode 00A0)
        which prevent wildcard find/replace patterns from matching date strings.

        Args:
            doc: The Word document object.
            allowed_story_types: Optional set of Word StoryType constants to
                restrict the scope of replacement.
        """
        try:
            for story in self._iter_story_ranges(
                doc, allowed_story_types=allowed_story_types
            ):
                f = story.Find
                f.ClearFormatting()
                f.Replacement.ClearFormatting()
                f.Execute(
                    "^s",  # FindText: ^s = non-breaking space in Word
                    False,  # MatchCase
                    False,  # MatchWholeWord
                    False,  # MatchWildcards (must be False for special codes)
                    False,  # MatchSoundsLike
                    False,  # MatchAllWordForms
                    True,  # Forward
                    WD_FIND_CONTINUE,  # Wrap
                    False,  # Format
                    " ",  # ReplaceWith: regular space
                    WD_REPLACE_ALL,  # Replace
                )
        except Exception as e:
            logger.debug(f"Non-breaking space normalization: {e}")

    def _execute_replace(
        self,
        doc: Any,
        find_text: str,
        replace_text: str,
        allowed_story_types: Optional[set[int]] = None,
    ) -> bool:
        """
        Execute a find and replace operation across all story ranges.

        Args:
            doc: The Word document object
            find_text: The text pattern to find
            replace_text: The replacement text
            allowed_story_types: Optional set of Word StoryType constants to
                restrict which story ranges are searched.

        Returns:
            True if at least one replacement was made
        """
        any_replaced = False
        try:
            for story in self._iter_story_ranges(
                doc, allowed_story_types=allowed_story_types
            ):
                if self._run_find_replace(story, find_text, replace_text):
                    any_replaced = True
        except Exception as e:
            logger.warning(f"Error during find/replace: {e}")
        return any_replaced

    def _iter_story_ranges(
        self, doc: Any, allowed_story_types: Optional[set[int]] = None
    ) -> Iterator[Any]:
        """Iterate all story ranges, optionally filtering by StoryType.

        Args:
            doc: The Word document object.
            allowed_story_types: If provided, only yield ranges whose
                ``StoryType`` is in this set.

        Yields:
            Word Range objects from the document's StoryRanges collection.
        """

        try:
            for story in doc.StoryRanges:
                # Include this story and its linked NextStoryRange chain
                cur = story
                while cur:
                    stype = getattr(cur, "StoryType", None)
                    if allowed_story_types is None or stype in allowed_story_types:
                        yield cur
                    cur = getattr(cur, "NextStoryRange", None)
        except Exception as e:
            logger.debug(f"Error iterating story ranges: {e}")

    def _run_find_replace(
        self, range_obj: Any, find_text: str, replace_text: str
    ) -> bool:
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
                find_text,  # FindText
                False,  # MatchCase
                False,  # MatchWholeWord
                True,  # MatchWildcards
                False,  # MatchSoundsLike
                False,  # MatchAllWordForms
                True,  # Forward
                WD_FIND_CONTINUE,  # Wrap
                False,  # Format
                replace_text,  # ReplaceWith
                WD_REPLACE_ALL,  # Replace
            )
            if result:
                logger.debug(f"Find/replace matched: '{find_text}' -> '{replace_text}'")
            return bool(result)
        except Exception as e:
            logger.warning(f"Error in find/replace operation: {e}")
            return False

    def __del__(self) -> None:
        """Safety net: ensure COM resources are released if not explicitly shut down."""
        if self._initialized or self.word_app is not None:
            try:
                self.shutdown()
            except Exception as e:
                # Destructor — logging may itself be torn down; best-effort.
                try:
                    logger.debug(f"Error in __del__ cleanup: {e}")
                except Exception:
                    pass

    def __enter__(self) -> "WordProcessor":
        """Context manager entry: initialize Word and return self.

        Returns:
            The initialized ``WordProcessor`` instance.
        """
        self.initialize()
        return self

    def __exit__(
        self,
        exc_type: Optional[type[BaseException]],
        exc_val: Optional[BaseException],
        exc_tb: Optional[TracebackType],
    ) -> None:
        """Context manager exit: shut down the Word application.

        Args:
            exc_type: Exception type, if any.
            exc_val: Exception value, if any.
            exc_tb: Exception traceback, if any.
        """
        self.shutdown()
