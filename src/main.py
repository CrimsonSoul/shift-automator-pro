"""
Shift Automator - Main Application Entry Point

A high-performance desktop application for automating shift schedule printing.
"""

import sys
import threading
import time
from datetime import date, timedelta
from typing import Optional, Callable

import tkinter as tk

from .config import ConfigManager, AppConfig
from .constants import (
    PROGRESS_MAX,
    PRINT_MAX_RETRIES,
    PRINT_INITIAL_DELAY,
    PRINT_MAX_DELAY,
    TRANSIENT_ERROR_KEYWORDS,
    CONFIG_DEBOUNCE_DELAY,
    PRINTER_DEFAULT_PLACEHOLDER
)
from .logger import setup_logging, get_logger
from .path_validation import validate_folder_path
from .scheduler import get_shift_template_name, validate_date_range
from .ui import ScheduleAppUI
from .word_processor import WordProcessor, HAS_PYWIN32

# Set up logging
setup_logging()
logger = get_logger(__name__)


def _is_transient_error(error_message: str) -> bool:
    """
    Check if an error message indicates a transient failure that can be retried.

    Args:
        error_message: The error message to check

    Returns:
        True if the error appears to be transient, False otherwise
    """
    error_lower = error_message.lower()
    return any(keyword in error_lower for keyword in TRANSIENT_ERROR_KEYWORDS)


def _calculate_retry_delay(attempt: int, initial_delay: float, max_delay: float) -> float:
    """
    Calculate retry delay with exponential backoff.

    Args:
        attempt: The attempt number (0-indexed)
        initial_delay: Initial delay in seconds
        max_delay: Maximum delay in seconds

    Returns:
        Delay in seconds for this attempt
    """
    delay = initial_delay * (2 ** attempt)
    return min(delay, max_delay)


class ShiftAutomatorApp:
    """Main application controller."""

    def __init__(self, root: tk.Tk):
        """
        Initialize the application.

        Args:
            root: The Tkinter root window
        """
        self.root = root
        self.ui = ScheduleAppUI(root)
        self.config_manager = ConfigManager()
        self.word_processor: Optional[WordProcessor] = None
        self._processing_thread: Optional[threading.Thread] = None
        self._cancel_requested = False
        self._cancel_lock = threading.Lock()

        # Config debouncing
        self._config_save_pending = False
        self._config_save_timer: Optional[threading.Timer] = None
        self._config_save_lock = threading.Lock()

        # Load and apply saved configuration
        self._load_config()

        # Set up button commands
        self.ui.set_start_command(self.start_processing)
        self.ui.set_cancel_command(self.cancel_processing)

        # Set up config change callback to save configuration on changes
        self.ui.set_config_change_callback(self._schedule_config_save)

        logger.info("Shift Automator application initialized")

    def _load_config(self) -> None:
        """Load configuration and apply to UI."""
        try:
            config = self.config_manager.load()
            if config.day_folder and self.ui.day_entry:
                self.ui.day_entry.delete(0, tk.END)
                self.ui.day_entry.insert(0, config.day_folder)
            if config.night_folder and self.ui.night_entry:
                self.ui.night_entry.delete(0, tk.END)
                self.ui.night_entry.insert(0, config.night_folder)
            if config.printer_name and self.ui.printer_var:
                self.ui.printer_var.set(config.printer_name)
            logger.info("Configuration loaded successfully")
        except Exception as e:
            logger.error(f"Error loading configuration: {e}")
            self.ui.show_warning("Configuration Error",
                               f"Could not load saved configuration: {e}")

    def _save_config(self, config: AppConfig) -> None:
        """
        Save configuration.

        Args:
            config: Configuration to save
        """
        try:
            self.config_manager.save(config)
            logger.info("Configuration saved successfully")
        except Exception as e:
            logger.error(f"Error saving configuration: {e}")

    def _schedule_config_save(self) -> None:
        """Schedule a debounced configuration save."""
        with self._config_save_lock:
            # Cancel any pending save timer
            if self._config_save_timer is not None:
                self._config_save_timer.cancel()

            # Set flag that save is pending
            self._config_save_pending = True

            # Schedule new save
            self._config_save_timer = threading.Timer(
                CONFIG_DEBOUNCE_DELAY,
                self._save_config_if_pending
            )
            self._config_save_timer.start()

    def _save_config_if_pending(self) -> None:
        """Save configuration if a save is pending (called by timer)."""
        with self._config_save_lock:
            if self._config_save_pending:
                try:
                    self._save_current_config()
                except Exception as e:
                    logger.error(f"Failed to save config: {e}")
                    # Show error to user (schedule on main thread) if window still exists
                    try:
                        if self.root.winfo_exists():
                            self.root.after(0, lambda: self.ui.show_warning(
                                "Configuration Save Error",
                                f"Could not save configuration: {e}\n\nYour settings may not be preserved."
                            ))
                    except tk.TclError:
                        logger.warning("Cannot show config save error - window already destroyed")
                    except Exception as ui_error:
                        logger.error(f"Could not show config save error to user: {ui_error}")
                finally:
                    # Always clear the pending flags to allow future saves
                    self._config_save_pending = False
                    self._config_save_timer = None

    def _save_current_config(self) -> None:
        """Save the current configuration from UI to file."""
        config = AppConfig(
            day_folder=self.ui.get_day_folder(),
            night_folder=self.ui.get_night_folder(),
            printer_name=self.ui.get_printer_name()
        )
        self._save_config(config)

    def _validate_inputs(self) -> tuple[bool, Optional[str]]:
        """
        Validate all user inputs before processing.

        Returns:
            Tuple of (is_valid, error_message)
        """
        day_folder = self.ui.get_day_folder()
        night_folder = self.ui.get_night_folder()
        printer_name = self.ui.get_printer_name()

        # Check that folders are provided
        if not day_folder:
            return False, "Please select a Day Templates folder"
        if not night_folder:
            return False, "Please select a Night Templates folder"
        if not printer_name or printer_name == PRINTER_DEFAULT_PLACEHOLDER:
            return False, "Please select a target printer"

        # Validate folder paths
        is_valid, error_msg = validate_folder_path(day_folder)
        if not is_valid:
            return False, f"Invalid Day Templates folder: {error_msg}"

        is_valid, error_msg = validate_folder_path(night_folder)
        if not is_valid:
            return False, f"Invalid Night Templates folder: {error_msg}"

        # Validate date range
        start_date = self.ui.get_start_date()
        end_date = self.ui.get_end_date()

        if not start_date or not end_date:
            return False, "Please select both start and end dates"

        is_valid, error_msg = validate_date_range(start_date, end_date)
        if not is_valid:
            return False, error_msg

        return True, None

    def start_processing(self) -> None:
        """Start the batch processing in a background thread."""
        self.ui.log("Validating configuration...")
        
        # Validate inputs
        is_valid, error_msg = self._validate_inputs()
        if not is_valid:
            self.ui.log(f"Validation failed: {error_msg}")
            self.ui.show_warning("Validation Error", error_msg or "Unknown error")
            return

        # Reset cancellation flag
        self._cancel_requested = False

        # Update UI for processing state
        self.ui.set_print_button_state("disabled")
        self.ui.set_cancel_button_state("normal")
        self.ui.log("Starting background process...")

        # Start processing thread
        self._processing_thread = threading.Thread(
            target=self._process_batch,
            daemon=True
        )
        self._processing_thread.start()

    def cancel_processing(self) -> None:
        """Request cancellation of the current batch processing."""
        if self._processing_thread and self._processing_thread.is_alive():
            with self._cancel_lock:
                self._cancel_requested = True
            self.ui.log("Cancellation requested, finishing current document...")
            self.ui.set_cancel_button_state("disabled")
            logger.info("User requested cancellation of batch processing")

    def _print_with_retry(self, word_proc: WordProcessor, folder: str, template: str,
                          current_date: date, printer_name: str) -> tuple[bool, Optional[str]]:
        """
        Print a document with retry logic for transient failures.

        Args:
            word_proc: The WordProcessor instance
            folder: Template folder path
            template: Template name
            current_date: Date for the document
            printer_name: Target printer name

        Returns:
            Tuple of (success, error_message)
        """
        last_error = None

        for attempt in range(PRINT_MAX_RETRIES):
            # Check for cancellation before retry
            with self._cancel_lock:
                if self._cancel_requested:
                    logger.info(f"Print cancelled during retry attempt {attempt + 1}")
                    return False, "Cancelled by user"

            # Attempt to print
            success, error = word_proc.print_document(
                folder, template, current_date, printer_name
            )

            if success:
                if attempt > 0:
                    logger.info(f"Successfully printed {template} after {attempt} retries")
                return True, None

            last_error = error

            # Check if error is transient and we should retry
            if attempt < PRINT_MAX_RETRIES - 1 and _is_transient_error(error or ""):
                delay = _calculate_retry_delay(attempt, PRINT_INITIAL_DELAY, PRINT_MAX_DELAY)
                logger.warning(
                    f"Transient error printing {template} (attempt {attempt + 1}/{PRINT_MAX_RETRIES}): "
                    f"{error}. Retrying in {delay:.1f} seconds..."
                )
                time.sleep(delay)
            else:
                # Not a transient error or no more retries
                break

        return False, last_error

    def _schedule_ui_update(self, callback: Callable, *args) -> None:
        """
        Schedule a UI callback on the main thread.

        This replaces lambda capture pattern with a clearer helper method.

        Args:
            callback: The UI method to call
            *args: Arguments to pass to the callback
        """
        def _update():
            callback(*args)
        self.root.after(0, _update)

    def _schedule_log(self, message: str, progress: Optional[float] = None) -> None:
        """
        Schedule a log update on the main thread.

        Args:
            message: Message to log
            progress: Optional progress value to update
        """
        def _update():
            self.ui.log(message)
            if progress is not None:
                self.ui.update_progress(progress)
        self.root.after(0, _update)

    def _process_batch(self) -> None:
        """Process the batch of schedules."""
        start_date = self.ui.get_start_date()
        end_date = self.ui.get_end_date()
        day_folder = self.ui.get_day_folder()
        night_folder = self.ui.get_night_folder()
        printer_name = self.ui.get_printer_name()

        # Check if dates are None (should be handled by validation, but for type safety)
        if not start_date or not end_date:
            logger.error("Start date or end date is None in _process_batch")
            return

        # Cancel any pending config save and save immediately before processing
        with self._config_save_lock:
            if self._config_save_timer is not None:
                self._config_save_timer.cancel()
                self._config_save_timer = None
            self._config_save_pending = False

        # Save configuration
        config = AppConfig(
            day_folder=day_folder,
            night_folder=night_folder,
            printer_name=printer_name
        )
        self._save_config(config)

        # Calculate total days
        total_days = (end_date - start_date).days + 1
        logger.info(f"Processing {total_days} days from {start_date} to {end_date}")

        # Track failed operations
        failed_operations = []

        cancelled = False
        # processed_days tracking: counts fully or partially completed days
        processed_days = 0
        progress = 0.0  # Initialize progress for early cancellation case
        try:
            self.root.after(0, lambda: self.ui.log("Initializing Word Processor (this may take a moment)..."))
            with WordProcessor() as word_proc:
                self.root.after(0, lambda: self.ui.log("Word connection established."))
                for i in range(total_days):
                    # Check for cancellation request before processing this day
                    with self._cancel_lock:
                        if self._cancel_requested:
                            cancelled = True
                            processed_days = i  # No work done on day i yet
                            logger.info(f"Processing cancelled by user after {i} of {total_days} days")
                            break

                    current_date = start_date + timedelta(days=i)
                    day_name = current_date.strftime("%A")
                    display_date = current_date.strftime("%m/%d/%Y")

                    # Update progress
                    progress = (i / total_days) * 100
                    self._schedule_log(f"Processing {day_name} {display_date}...", progress)

                    # Process Day Shift
                    day_template = get_shift_template_name(current_date, "day")
                    success, error = self._print_with_retry(
                        word_proc, day_folder, day_template, current_date, printer_name
                    )
                    if not success:
                        failed_operations.append({
                            'date': current_date,
                            'shift': 'day',
                            'template': day_template,
                            'error': error
                        })
                        self._schedule_log(f"Error printing day shift: {error}")
                        logger.error(f"Failed to print day shift for {current_date}: {error}")
                    else:
                        self._schedule_log(f"Sent {day_name} Day Shift to printer.")

                    # Check for cancellation request before night shift
                    with self._cancel_lock:
                        if self._cancel_requested:
                            cancelled = True
                            processed_days = i + 1  # Day shift completed, so day i is partially done
                            logger.info(f"Processing cancelled by user after day shift of {current_date}")
                            break

                    # Process Night Shift
                    night_template = get_shift_template_name(current_date, "night")
                    success, error = self._print_with_retry(
                        word_proc, night_folder, night_template, current_date, printer_name
                    )
                    if not success:
                        failed_operations.append({
                            'date': current_date,
                            'shift': 'night',
                            'template': night_template,
                            'error': error
                        })
                        self._schedule_log(f"Error printing night shift: {error}")
                        logger.error(f"Failed to print night shift for {current_date}: {error}")
                    else:
                        self._schedule_log(f"Sent {day_name} Night Shift to printer.")

                # Show completion status
                if cancelled:
                    self._schedule_log("Processing CANCELLED by user.")
                    self._schedule_ui_update(self.ui.show_info,
                        "Processing Cancelled",
                        f"Cancelled by user.\n\nProcessed {processed_days} of {total_days} days before cancellation.")
                else:
                    self._schedule_log("Processing COMPLETE!")
                    self._schedule_ui_update(self.ui.update_progress, PROGRESS_MAX)

                    # Show results
                    if failed_operations:
                        self._show_failure_summary(failed_operations)
                    else:
                        self.root.after(0, lambda: self.ui.show_info(
                            "Success",
                            f"All {total_days} days have been processed and sent to the printer."
                        ))

        except Exception as e:
            # Capture full exception details
            error_type = type(e).__name__
            error_msg = str(e)
            full_error = f"{error_type}: {error_msg}"

            logger.exception("Error during batch processing")
            self._schedule_log(f"FATAL ERROR: {full_error}")
            self._schedule_ui_update(self.ui.show_error,
                "Processing Error",
                f"An error occurred during processing:\n\n{full_error}")
        finally:
            # Re-enable buttons
            self._schedule_ui_update(self.ui.set_print_button_state, "normal")
            self._schedule_ui_update(self.ui.set_cancel_button_state, "disabled")

    def _show_failure_summary(self, failed_operations: list[dict]) -> None:
        """
        Show a summary of failed operations.

        Args:
            failed_operations: List of failed operation details
        """
        total = len(failed_operations)
        message = f"{total} operation(s) failed:\n\n"

        # Show first 5 failures
        for i, op in enumerate(failed_operations[:5], 1):
            date_str = op['date'].strftime("%m/%d/%Y")
            message += f"{i}. {date_str} {op['shift'].title()} Shift ({op['template']}): {op['error']}\n"

        if total > 5:
            message += f"\n... and {total - 5} more failures"

        message += "\n\nCheck the log file for details."

        self.ui.show_warning("Processing Completed with Errors", message)


def main() -> None:
    """Main entry point for the application."""
    logger.info("Starting Shift Automator")

    # Early check for pywin32 availability (Windows-only)
    if sys.platform == "win32" and not HAS_PYWIN32:
        error_msg = (
            "This application requires pywin32 to be installed.\n\n"
            "The pywin32 package is missing or could not be loaded.\n"
            "Please reinstall the application or contact support."
        )
        logger.error(f"pywin32 not available: HAS_PYWIN32={HAS_PYWIN32}")
        try:
            import tkinter.messagebox as mb
            # Need a root window for messagebox
            temp_root = tk.Tk()
            temp_root.withdraw()
            mb.showerror("Missing Dependency", error_msg)
            temp_root.destroy()
        except Exception:
            print(f"FATAL: {error_msg}")
        return

    try:
        root = tk.Tk()
        app = ShiftAutomatorApp(root)
        app.ui.run()
    except Exception as e:
        logger.exception("Fatal error in main")
        try:
            import tkinter.messagebox as mb
            mb.showerror("Fatal Error", f"The application encountered a fatal error:\n\n{str(e)}")
        except Exception as messagebox_error:
            print(f"Fatal error: {e}")
            print(f"Error showing message box: {messagebox_error}")
    finally:
        logger.info("Shift Automator shutting down")


if __name__ == "__main__":
    main()
