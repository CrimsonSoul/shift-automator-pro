"""
Shift Automator - Main Application Entry Point

A high-performance desktop application for automating shift schedule printing.
"""

import threading
from datetime import date, timedelta
from typing import Optional

import tkinter as tk

from .config import ConfigManager, AppConfig
from .constants import PROGRESS_MAX, COLORS

from .logger import setup_logging, get_logger
from .path_validation import validate_folder_path
from .scheduler import get_shift_template_name, validate_date_range
from .ui import ScheduleAppUI
from .word_processor import WordProcessor

# Set up logging
setup_logging()
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
        self._cancel_requested = False


        # Load and apply saved configuration
        self._load_config()

        # Set up button command
        self.ui.set_start_command(self.start_processing)

        logger.info("Shift Automator application initialized")

    def _load_config(self) -> None:
        """Load configuration and apply to UI."""
        try:
            config = self.config_manager.load()
            if config.day_folder:
                self.ui.day_entry.delete(0, tk.END)
                self.ui.day_entry.insert(0, config.day_folder)
            if config.night_folder:
                self.ui.night_entry.delete(0, tk.END)
                self.ui.night_entry.insert(0, config.night_folder)
            if config.printer_name:
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
        if not printer_name or printer_name == "Choose Printer":
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
        # Check if already processing (can be used as stop button)
        if self._processing_thread and self._processing_thread.is_alive():
            self._cancel_requested = True
            self.ui.update_status("Cancelling...", self.ui.progress_var.get())
            self.ui.set_print_button_state("disabled")
            return

        # Validate inputs
        is_valid, error_msg = self._validate_inputs()
        if not is_valid:
            self.ui.show_warning("Validation Error", error_msg)
            return

        # Reset cancel flag
        self._cancel_requested = False

        # Update button text to STOP
        self.ui.print_btn.config(text="STOP EXECUTION", bg=COLORS.success) # Use success color for stop? No, maybe red? 
        # Actually I'll use COLORS.accent for both or change it.
        
        # Start processing thread
        self._processing_thread = threading.Thread(
            target=self._process_batch,
            daemon=True
        )
        self._processing_thread.start()

    def _process_batch(self) -> None:
        """Process the batch of schedules."""
        start_date = self.ui.get_start_date()
        end_date = self.ui.get_end_date()
        
        if not start_date or not end_date:
            logger.error("Attempted to process batch with missing dates")
            return

        day_folder = self.ui.get_day_folder()
        night_folder = self.ui.get_night_folder()
        printer_name = self.ui.get_printer_name()


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

        try:
            with WordProcessor() as word_proc:
                for i in range(total_days):
                    # Check for cancellation
                    if self._cancel_requested:
                        logger.info("Batch processing cancelled by user")
                        self.root.after(0, lambda: self.ui.update_status("Cancelled", 0))
                        return

                    current_date = start_date + timedelta(days=i)

                    day_name = current_date.strftime("%A")
                    display_date = current_date.strftime("%m/%d/%Y")

                    # Update progress
                    progress = (i / total_days) * 100
                    self.root.after(0, lambda m=f"Processing {day_name} {display_date}...",
                                   p=progress: self.ui.update_status(m, p))

                    # Process Day Shift
                    day_template = get_shift_template_name(current_date, "day")
                    success, error = word_proc.print_document(
                        day_folder, day_template, current_date, printer_name
                    )
                    if not success:
                        failed_operations.append({
                            'date': current_date,
                            'shift': 'day',
                            'template': day_template,
                            'error': error
                        })
                        logger.error(f"Failed to print day shift for {current_date}: {error}")

                    # Process Night Shift
                    night_template = get_shift_template_name(current_date, "night")
                    success, error = word_proc.print_document(
                        night_folder, night_template, current_date, printer_name
                    )
                    if not success:
                        failed_operations.append({
                            'date': current_date,
                            'shift': 'night',
                            'template': night_template,
                            'error': error
                        })
                        logger.error(f"Failed to print night shift for {current_date}: {error}")

                # Complete
                self.root.after(0, lambda: self.ui.update_status("Complete!", PROGRESS_MAX))

                # Show results
                if failed_operations:
                    self._show_failure_summary(failed_operations)
                else:
                    self.root.after(0, lambda: self.ui.show_info(
                        "Success",
                        f"All {total_days} days have been processed and sent to the printer."
                    ))

        except Exception as e:
            logger.exception("Error during batch processing")
            self.root.after(0, lambda: self.ui.show_error(
                "Processing Error",
                f"An error occurred during processing: {str(e)}"
            ))
        finally:
            # Re-enable button and reset text
            self.root.after(0, lambda: self.ui.print_btn.config(text="START EXECUTION", bg=COLORS.accent))
            self.root.after(0, lambda: self.ui.set_print_button_state("normal"))


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

    try:
        root = tk.Tk()
        app = ShiftAutomatorApp(root)
        app.ui.run()
    except Exception as e:
        logger.exception("Fatal error in main")
        try:
            import tkinter.messagebox as mb
            mb.showerror("Fatal Error", f"The application encountered a fatal error:\n\n{str(e)}")
        except Exception:
            print(f"Fatal error: {e}")
    finally:
        logger.info("Shift Automator shutting down")



if __name__ == "__main__":
    main()
