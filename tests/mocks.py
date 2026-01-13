import os
import threading
from datetime import date
from typing import Optional, List, Tuple, Dict, Any
from unittest.mock import MagicMock

class MockWordProcessor:
    """Mock implementation of WordProcessor for E2E testing."""
    def __init__(self):
        self.initialized = False
        self.print_log: List[Dict[str, Any]] = []
        self.should_fail = False
        self.error_message = "Mock Error"
        self._thread_id = None
        self.call_count = 0

    def initialize(self) -> None:
        self.initialized = True
        self._thread_id = threading.get_ident()

    def shutdown(self) -> None:
        self.initialized = False

    def print_document(self, folder: str, template_name: str, current_date: date,
                       printer_name: str) -> Tuple[bool, Optional[str]]:
        self.call_count += 1
        if not self.initialized:
            return False, "Not initialized"
        
        if self.should_fail:
            return False, self.error_message

        # Record the call
        self.print_log.append({
            "folder": folder,
            "template": template_name,
            "date": current_date,
            "printer": printer_name
        })
        return True, None

    def __enter__(self):
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.shutdown()

class MockUI:
    """Mock UI to capture updates and simulate user input."""
    def __init__(self, root=None):
        self.root = root or MagicMock()
        self.status_messages: List[str] = []
        self.progress_values: List[float] = []
        self.error_calls: List[Tuple[str, str]] = []
        self.info_calls: List[Tuple[str, str]] = []
        self.warning_calls: List[Tuple[str, str]] = []
        
        self.day_folder = ""
        self.night_folder = ""
        self.printer_name = "Mock Printer"
        self.start_date = date(2026, 1, 1)
        self.end_date = date(2026, 1, 7)
        
        self.start_command = None
        self.cancel_command = None
        
        # Mock widgets to avoid AttributeError
        self.day_entry = MagicMock()
        self.night_entry = MagicMock()
        self.printer_var = MagicMock()
        self.printer_dropdown = MagicMock()
        self.status_label = MagicMock()
        self.progress = MagicMock()
        self.print_btn = MagicMock()
        self.cancel_btn = MagicMock()
        
        # Mock methods that tests want to verify
        self.show_warning = MagicMock(side_effect=self._record_warning)
        self.show_error = MagicMock(side_effect=self._record_error)
        self.show_info = MagicMock(side_effect=self._record_info)
        self.update_status = MagicMock(side_effect=self._record_status)

    @property
    def status_history(self) -> List[str]:
        return self.status_messages

    @property
    def info_messages(self) -> List[str]:
        return [call[1] for call in self.info_calls]

    @property
    def warning_messages(self) -> List[str]:
        return [call[1] for call in self.warning_calls]

    def _record_status(self, message: str, progress: Optional[float]) -> None:
        self.status_messages.append(message)
        if progress is not None:
            self.progress_values.append(progress)

    def _record_error(self, title: str, message: str) -> None:
        self.error_calls.append((title, message))

    def _record_warning(self, title: str, message: str) -> None:
        self.warning_calls.append((title, message))
        # Also record warnings to error_calls for easier checking in some tests
        self.error_calls.append((title, message))

    def _record_info(self, title: str, message: str) -> None:
        self.info_calls.append((title, message))

    def get_day_folder(self) -> str: return self.day_folder
    def get_night_folder(self) -> str: return self.night_folder
    def get_printer_name(self) -> str: return self.printer_name
    def get_start_date(self) -> date: return self.start_date
    def get_end_date(self) -> date: return self.end_date

    def set_dates(self, start: date, end: date) -> None:
        self.start_date = start
        self.end_date = end

    def set_folders(self, day: str, night: str) -> None:
        self.day_folder = day
        self.night_folder = night

    def set_printer(self, printer: str) -> None:
        self.printer_name = printer

    def set_start_command(self, cmd): self.start_command = cmd
    def set_cancel_command(self, cmd): self.cancel_command = cmd
    
    def set_print_button_state(self, state): pass
    def set_cancel_button_state(self, state): pass
    def set_config_change_callback(self, cb): pass
