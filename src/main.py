"""
Shift Automator - Main Application Entry Point

A high-performance desktop application for automating shift schedule printing.
"""

import threading
import csv
from datetime import datetime
from datetime import timedelta
from typing import Optional, Callable

import tkinter as tk

from .config import ConfigManager, AppConfig
from .constants import PROGRESS_MAX, COLORS, DEFAULT_PRINTER_LABEL, LOG_FILENAME

from .logger import setup_logging, get_logger
from .path_validation import validate_folder_path
from .scheduler import get_shift_template_name, validate_date_range
from .ui import ScheduleAppUI
from .word_processor import (
    WordProcessor,
    TemplateLookupError,
    get_word_automation_status,
)
from .app_paths import get_data_dir

logger = get_logger(__name__)


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
        self._cancel_event = threading.Event()
        self._closing = False

        # Load and apply saved configuration
        self._load_config()

        # Set up button command
        self.ui.set_start_command(self.start_processing)

        # Handle window close gracefully
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        logger.info("Shift Automator application initialized")

    def _safe_after(self, callback: Callable[[], None]) -> None:
        """Schedule a UI callback if the window is still alive.

        Tkinter isn't thread-safe; all UI updates must be scheduled onto the UI thread.
        During shutdown, the root window can be destroyed while the worker thread is
        still running; in that case, after() can raise TclError.
        """

        if self._closing:
            return

        try:
            self.root.after(0, callback)
        except tk.TclError:
            # Window already destroyed; ignore late UI updates.
            logger.debug("UI update skipped (window closed)")

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
            if self.ui.headers_only_var:
                self.ui.headers_only_var.set(
                    bool(getattr(config, "headers_footers_only", False))
                )
            logger.info("Configuration loaded successfully")
        except Exception as e:
            logger.error(f"Error loading configuration: {e}")
            self.ui.show_warning(
                "Configuration Error", f"Could not load saved configuration: {e}"
            )

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

    def _validate_inputs(self) -> tuple[bool, Optional[str]]:
        """
        Validate all user inputs before processing.

        Returns:
            Tuple of (is_valid, error_message)
        """
        day_folder = (self.ui.get_day_folder() or "").strip()
        night_folder = (self.ui.get_night_folder() or "").strip()
        printer_name = (self.ui.get_printer_name() or "").strip()
        start_date = self.ui.get_start_date()
        end_date = self.ui.get_end_date()

        # Check that folders are provided
        if not day_folder:
            return False, "Please select a Day Templates folder"
        if not night_folder:
            return False, "Please select a Night Templates folder"
        if not printer_name or printer_name == DEFAULT_PRINTER_LABEL:
            return False, "Please select a target printer"

        # Validate printer against enumerated list when available
        available_printers = []
        try:
            available_printers = self.ui.get_available_printers()
        except Exception:
            available_printers = []

        if available_printers and printer_name not in available_printers:
            return False, (
                "Selected printer is not available. Click Refresh and select a valid printer."
            )

        # Validate date range
        if not start_date or not end_date:
            return False, "Please select both start and end dates"

        is_valid, error_msg = validate_date_range(start_date, end_date)
        if not is_valid:
            return False, error_msg

        # Ensure Word automation bindings exist (pywin32)
        ok, word_err = get_word_automation_status()
        if not ok:
            return False, word_err

        # Validate folder paths
        is_valid, error_msg = validate_folder_path(day_folder)
        if not is_valid:
            return False, f"Invalid Day Templates folder: {error_msg}"

        is_valid, error_msg = validate_folder_path(night_folder)
        if not is_valid:
            return False, f"Invalid Night Templates folder: {error_msg}"

        # Preflight template availability (fail fast before opening Word/printing)
        ok, preflight_err = self._preflight_templates(
            day_folder, night_folder, start_date, end_date
        )
        if not ok:
            return False, preflight_err

        return True, None

    def _preflight_templates(
        self,
        day_folder: str,
        night_folder: str,
        start_date,
        end_date,
    ) -> tuple[bool, Optional[str]]:
        """Validate that all required templates exist and resolve unambiguously."""

        try:
            total_days = (end_date - start_date).days + 1
        except Exception:
            return False, "Invalid date range"

        day_templates: set[str] = set()
        night_templates: set[str] = set()

        for i in range(total_days):
            dt = start_date + timedelta(days=i)
            day_templates.add(get_shift_template_name(dt, "day"))
            night_templates.add(get_shift_template_name(dt, "night"))

        wp = WordProcessor()
        missing: list[str] = []

        def check(folder: str, templates: set[str], label: str) -> Optional[str]:
            for name in sorted(templates):
                try:
                    found = wp.find_template_file(folder, name)
                except TemplateLookupError as e:
                    return f"{label} template lookup error for '{name}': {e}"
                if not found:
                    missing.append(f"{label}: {name}")
            return None

        err = check(day_folder, day_templates, "Day")
        if err:
            return False, err

        err = check(night_folder, night_templates, "Night")
        if err:
            return False, err

        if missing:
            shown = "\n".join(missing[:10])
            more = ""
            if len(missing) > 10:
                more = f"\n...and {len(missing) - 10} more"
            return (
                False,
                "Missing required templates:\n\n"
                f"{shown}{more}\n\n"
                "Verify your template folders and naming conventions.",
            )

        return True, None

    def start_processing(self) -> None:
        """Start the batch processing in a background thread."""
        # Check if already processing (can be used as stop button)
        if self._processing_thread and self._processing_thread.is_alive():
            self._cancel_event.set()
            current_progress = (
                self.ui.progress_var.get() if self.ui.progress_var else 0.0
            )
            self.ui.update_status(
                "Stopping after current document...", current_progress
            )
            self.ui.set_print_button_state("disabled")
            return

        # Validate inputs
        is_valid, error_msg = self._validate_inputs()
        if not is_valid:
            self.ui.show_warning("Validation Error", error_msg or "Unknown error")
            return

        # Confirm large batches
        start_date = self.ui.get_start_date()
        end_date = self.ui.get_end_date()
        if start_date and end_date:
            total_days = (end_date - start_date).days + 1
            total_docs = total_days * 2
            if total_days >= 30:
                ok = self.ui.ask_yes_no(
                    "Large Batch Confirm",
                    f"This will print {total_docs} documents ({total_days} days x 2 shifts).\n\nContinue?",
                )
                if not ok:
                    self.ui.update_status("Cancelled by user", 0)
                    return

        # Reset cancel flag
        self._cancel_event.clear()

        # Collect all UI values in the main thread (Tkinter is not thread-safe)
        batch_params = {
            "start_date": self.ui.get_start_date(),
            "end_date": self.ui.get_end_date(),
            "day_folder": self.ui.get_day_folder(),
            "night_folder": self.ui.get_night_folder(),
            "printer_name": self.ui.get_printer_name(),
            "headers_footers_only": self.ui.get_headers_footers_only(),
        }

        # Update button text to STOP
        if self.ui.print_btn:
            self.ui.print_btn.config(text="STOP EXECUTION", bg=COLORS.error)

        # Start processing thread with pre-collected values
        self._processing_thread = threading.Thread(
            target=self._process_batch, args=(batch_params,), daemon=True
        )
        self._processing_thread.start()

    def _process_batch(self, params: dict) -> None:
        """
        Process the batch of schedules.

        Args:
            params: Pre-collected UI values with keys: start_date, end_date,
                    day_folder, night_folder, printer_name
        """
        start_date = params["start_date"]
        end_date = params["end_date"]

        if not start_date or not end_date:
            logger.error("Attempted to process batch with missing dates")
            return

        day_folder = params["day_folder"]
        night_folder = params["night_folder"]
        printer_name = params["printer_name"]
        headers_footers_only = bool(params.get("headers_footers_only", False))

        # Save configuration
        config = AppConfig(
            day_folder=day_folder,
            night_folder=night_folder,
            printer_name=printer_name,
            headers_footers_only=headers_footers_only,
        )
        self._save_config(config)

        # Calculate total days (MAX_DAYS_RANGE already validated by _validate_inputs)
        total_days = (end_date - start_date).days + 1
        total_jobs = total_days * 2

        logger.info(f"Processing {total_days} days from {start_date} to {end_date}")

        # Track failed operations
        failed_operations = []

        try:
            with WordProcessor() as word_proc:
                job_index = 0
                for i in range(total_days):
                    # Check for cancellation
                    if self._cancel_event.is_set():
                        logger.info("Batch processing cancelled by user")
                        self._safe_after(lambda: self.ui.update_status("Cancelled", 0))
                        return

                    current_date = start_date + timedelta(days=i)

                    day_name = current_date.strftime("%A")
                    display_date = current_date.strftime("%m/%d/%Y")

                    # Update progress (per document)
                    progress = (job_index / max(total_jobs, 1)) * 100
                    prep_msg = f"Preparing {day_name} {display_date}..."
                    prep_progress = progress

                    def _update_preparing(
                        msg: str = prep_msg,
                        prog: float = prep_progress,
                    ) -> None:
                        self.ui.update_status(msg, prog)

                    self._safe_after(_update_preparing)

                    # Process Day Shift
                    day_template = get_shift_template_name(current_date, "day")
                    progress = (job_index / max(total_jobs, 1)) * 100
                    day_msg = (
                        f"Printing Day Shift: {day_name} {display_date} "
                        f"({job_index + 1}/{total_jobs})..."
                    )
                    day_progress = progress

                    def _update_day(
                        msg: str = day_msg,
                        prog: float = day_progress,
                    ) -> None:
                        self.ui.update_status(msg, prog)

                    self._safe_after(_update_day)

                    if self._cancel_event.is_set():
                        logger.info("Batch processing cancelled by user")
                        self._safe_after(lambda: self.ui.update_status("Cancelled", 0))
                        return

                    success, error = word_proc.print_document(
                        day_folder,
                        day_template,
                        current_date,
                        printer_name,
                        headers_footers_only=headers_footers_only,
                    )
                    job_index += 1
                    if not success:
                        failed_operations.append(
                            {
                                "date": current_date,
                                "shift": "day",
                                "template": day_template,
                                "error": error,
                            }
                        )
                        logger.error(
                            f"Failed to print day shift for {current_date}: {error}"
                        )

                    # Process Night Shift
                    night_template = get_shift_template_name(current_date, "night")
                    progress = (job_index / max(total_jobs, 1)) * 100
                    night_msg = (
                        f"Printing Night Shift: {day_name} {display_date} "
                        f"({job_index + 1}/{total_jobs})..."
                    )
                    night_progress = progress

                    def _update_night(
                        msg: str = night_msg,
                        prog: float = night_progress,
                    ) -> None:
                        self.ui.update_status(msg, prog)

                    self._safe_after(_update_night)

                    if self._cancel_event.is_set():
                        logger.info("Batch processing cancelled by user")
                        self._safe_after(lambda: self.ui.update_status("Cancelled", 0))
                        return

                    success, error = word_proc.print_document(
                        night_folder,
                        night_template,
                        current_date,
                        printer_name,
                        headers_footers_only=headers_footers_only,
                    )
                    job_index += 1
                    if not success:
                        failed_operations.append(
                            {
                                "date": current_date,
                                "shift": "night",
                                "template": night_template,
                                "error": error,
                            }
                        )
                        logger.error(
                            f"Failed to print night shift for {current_date}: {error}"
                        )

                # Complete
                self._safe_after(
                    lambda: self.ui.update_status("Complete!", PROGRESS_MAX)
                )

                # Show results
                if failed_operations:
                    self._safe_after(
                        lambda: self._show_failure_summary(failed_operations)
                    )
                else:
                    self._safe_after(
                        lambda: self.ui.show_info(
                            "Success",
                            f"All {total_days} days have been processed and sent to the printer.",
                        )
                    )

        except Exception as e:
            logger.exception("Error during batch processing")
            self._safe_after(
                lambda: self.ui.show_error(
                    "Processing Error", f"An error occurred during processing: {str(e)}"
                )
            )
        finally:
            # Re-enable button and reset text
            def reset_ui():
                if self._closing:
                    return
                if self.ui.print_btn:
                    self.ui.print_btn.config(text="START EXECUTION", bg=COLORS.accent)
                self.ui.set_print_button_state("normal")

            self._safe_after(reset_ui)

    def _on_close(self) -> None:
        """Handle window close: cancel any running batch and shut down cleanly."""
        self._closing = True
        if self._processing_thread and self._processing_thread.is_alive():
            logger.info("Window close requested during processing, cancelling...")
            self._cancel_event.set()
            self._processing_thread.join(timeout=5)
        self.root.destroy()

    def _show_failure_summary(self, failed_operations: list[dict]) -> None:
        """
        Show a summary of failed operations.

        Args:
            failed_operations: List of failed operation details
        """
        total = len(failed_operations)
        report_path = self._write_failure_report(failed_operations)
        message = f"{total} operation(s) failed:\n\n"

        # Show first 5 failures
        for i, op in enumerate(failed_operations[:5], 1):
            date_str = op["date"].strftime("%m/%d/%Y")
            message += f"{i}. {date_str} {op['shift'].title()} Shift ({op['template']}): {op['error']}\n"

        if total > 5:
            message += f"\n... and {total - 5} more failures"

        data_dir = get_data_dir()
        log_path = data_dir / LOG_FILENAME

        if report_path:
            message += f"\n\nFailure report saved to:\n{report_path}"
        message += f"\n\nLog file:\n{log_path}"
        message += "\n\nTip: Click 'Open Logs' in the app footer."

        self.ui.show_warning("Processing Completed with Errors", message)

    def _write_failure_report(self, failed_operations: list[dict]) -> Optional[str]:
        """Write a CSV failure report to the data directory."""

        try:
            data_dir = get_data_dir()
            data_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_file = data_dir / f"failure_report_{ts}.csv"

            with open(report_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["date", "shift", "template", "error"])
                for op in failed_operations:
                    writer.writerow(
                        [
                            op["date"].strftime("%m/%d/%Y"),
                            op["shift"],
                            op["template"],
                            op.get("error") or "",
                        ]
                    )
            logger.info(f"Failure report written: {report_file}")
            return str(report_file)
        except Exception as e:
            logger.warning(f"Could not write failure report: {e}")
            return None


def main() -> None:
    """Main entry point for the application."""
    setup_logging()
    logger.info("Starting Shift Automator")

    try:
        root = tk.Tk()
        app = ShiftAutomatorApp(root)
        app.ui.run()
    except Exception as e:
        logger.exception("Fatal error in main")
        try:
            import tkinter.messagebox as mb

            data_dir = get_data_dir()
            log_path = data_dir / LOG_FILENAME
            mb.showerror(
                "Fatal Error",
                "The application encountered a fatal error:\n\n"
                f"{str(e)}\n\n"
                f"Logs are saved to:\n{log_path}",
            )
        except Exception:
            print(f"Fatal error: {e}")
    finally:
        logger.info("Shift Automator shutting down")


if __name__ == "__main__":
    main()
