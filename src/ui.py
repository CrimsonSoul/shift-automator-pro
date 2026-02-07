"""
UI components for Shift Automator application.

This module contains all Tkinter UI components and styling.
"""

import os
import sys
import subprocess
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from datetime import date

try:
    from tkcalendar import DateEntry  # type: ignore
except Exception:  # pragma: no cover
    DateEntry = None

from typing import Optional, Callable, Literal, Union, Any, cast

try:
    import win32print  # type: ignore
except Exception:  # pragma: no cover
    win32print = None

from .constants import (
    COLORS,
    FONTS,
    WINDOW_WIDTH,
    WINDOW_HEIGHT,
    WINDOW_RESIZABLE,
    PROGRESS_MAX,
    PRINTER_ENUM_LOCAL,
    PRINTER_ENUM_NETWORK,
    DEFAULT_PRINTER_LABEL,
)
from .logger import get_logger
from .app_paths import get_data_dir

logger = get_logger(__name__)


class ScheduleAppUI:
    """Main UI class for the Shift Automator application."""

    def __init__(self, root: tk.Tk):
        """
        Initialize the UI.

        Args:
            root: The Tkinter root window
        """
        self.root = root
        self.root.title("Shift Automator")
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.resizable(WINDOW_RESIZABLE, WINDOW_RESIZABLE)
        self.root.configure(bg=COLORS.background)

        # Configure styles
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self._configure_styles()

        # UI components
        self.day_entry: Optional[ttk.Entry] = None
        self.night_entry: Optional[ttk.Entry] = None
        self.start_date_picker: Optional[Any] = None
        self.end_date_picker: Optional[Any] = None
        self.printer_var: Optional[tk.StringVar] = None
        self.headers_only_var: Optional[tk.BooleanVar] = None
        self.status_label: Optional[ttk.Label] = None
        self.progress_var: Optional[tk.DoubleVar] = None
        self.progress: Optional[ttk.Progressbar] = None
        self.printer_dropdown: Optional[ttk.OptionMenu] = None
        self.print_btn: Optional[tk.Button] = None

        # Cached enumerations
        self._cached_printers: list[str] = []

        # Create widgets
        self._create_widgets()

        logger.info("UI initialized")

    def _configure_styles(self) -> None:
        """Configure ttk styles for the application."""
        # Card Layout (Relay Style)
        self.style.configure("TFrame", background=COLORS.background)
        self.style.configure(
            "TLabel",
            background=COLORS.background,
            foreground=COLORS.text_main,
            font=FONTS.main,
        )
        self.style.configure(
            "TLabelframe",
            background=COLORS.background,
            foreground=COLORS.accent,
            bordercolor=COLORS.border,
            borderwidth=1,
        )
        self.style.configure(
            "TLabelframe.Label",
            background=COLORS.background,
            foreground=COLORS.text_dim,
            font=FONTS.sub,
        )

        # Inputs
        self.style.configure(
            "TEntry",
            fieldbackground=COLORS.surface,
            foreground=COLORS.text_main,
            insertcolor=COLORS.text_main,
            borderwidth=0,
        )

        # Buttons (Unified SaaS Look)
        self.style.configure(
            "TButton",
            background=COLORS.surface,
            foreground=COLORS.text_main,
            borderwidth=0,
            font=FONTS.bold,
            padding=(12, 6),
        )
        self.style.map("TButton", background=[("active", COLORS.secondary)])

        # Progress
        self.style.configure(
            "Horizontal.TProgressbar",
            thickness=4,
            troughcolor=COLORS.border,
            background=COLORS.accent,
        )

        # Specialized Labels
        self.style.configure(
            "Header.TLabel",
            font=FONTS.header,
            foreground=COLORS.text_main,
            background=COLORS.background,
        )
        self.style.configure(
            "Sub.TLabel",
            font=FONTS.sub,
            foreground=COLORS.text_dim,
            background=COLORS.background,
        )

        # Checkbuttons
        self.style.configure(
            "TCheckbutton",
            background=COLORS.background,
            foreground=COLORS.text_main,
            font=FONTS.sub,
        )

    def _enumerate_printers(self) -> list[str]:
        """Return a sorted list of available printer names."""

        if win32print is None:
            return []

        try:
            local_printers = [p[2] for p in win32print.EnumPrinters(PRINTER_ENUM_LOCAL)]
            network_printers = [
                p[2] for p in win32print.EnumPrinters(PRINTER_ENUM_NETWORK)
            ]
            return sorted(set(local_printers + network_printers))
        except Exception as e:
            logger.error(f"Error enumerating printers: {e}")
            return []

    def refresh_printers(self) -> None:
        """Re-enumerate printers and update the dropdown."""

        if not self.printer_dropdown or not self.printer_var:
            return

        printers = self._enumerate_printers()
        self._cached_printers = printers
        try:
            menu = self.printer_dropdown["menu"]
            menu.delete(0, "end")
            for name in printers:
                # tk._setit is an internal helper; keep usage isolated here.
                set_it = getattr(tk, "_setit", None)
                if callable(set_it):
                    menu.add_command(label=name, command=set_it(self.printer_var, name))
                else:
                    var = self.printer_var
                    if var is not None:
                        menu.add_command(
                            label=name, command=lambda v=name, vv=var: vv.set(v)
                        )

            current = self.printer_var.get()
            if current and current in printers:
                self.printer_var.set(current)
            else:
                self.printer_var.set(DEFAULT_PRINTER_LABEL)
        except Exception as e:
            logger.error(f"Could not update printer dropdown: {e}")

    def _create_widgets(self) -> None:
        """Create all UI widgets."""
        # Background Canvas
        bg_canvas = ttk.Frame(self.root, padding="40")
        bg_canvas.pack(fill="both", expand=True)

        # Header Section
        self._create_header(bg_canvas)

        # Config Card
        self._create_config_card(bg_canvas)

        # Control Card
        self._create_control_card(bg_canvas)

        # Action Footer
        self._create_footer(bg_canvas)

    def _create_header(self, parent: ttk.Frame) -> None:
        """Create the header section."""
        header_row = ttk.Frame(parent)
        header_row.pack(fill="x", pady=(0, 40))
        ttk.Label(header_row, text="Shift Automator", style="Header.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            header_row,
            text="High-performance batch scheduling & printing",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

    def _create_config_card(self, parent: ttk.Frame) -> None:
        """Create the configuration card."""
        config_card = ttk.LabelFrame(parent, text=" CONFIGURATION ", padding="24")
        config_card.pack(fill="x", pady=(0, 20))

        self.day_entry = self._create_path_row(config_card, "Day Templates", "")
        self.night_entry = self._create_path_row(config_card, "Night Templates", "")

    def _create_control_card(self, parent: ttk.Frame) -> None:
        """Create the controls card."""
        control_card = ttk.LabelFrame(parent, text=" CONTROLS ", padding="24")
        control_card.pack(fill="x")

        # Date Range Row
        self._create_date_range_row(control_card)

        # Printer Selection Row
        self._create_printer_row(control_card)

        # Advanced options
        self._create_options_row(control_card)

    def _create_date_range_row(self, parent: Union[ttk.Frame, ttk.LabelFrame]) -> None:
        """Create the date range selection row."""
        if DateEntry is None:
            raise RuntimeError(
                "Missing dependency: tkcalendar. Please reinstall requirements.txt and try again."
            )

        date_entry_cls = cast(Any, DateEntry)

        range_row = ttk.Frame(parent)
        range_row.pack(fill="x", pady=(0, 20))

        # Start Date
        start_wrap = ttk.Frame(range_row)
        start_wrap.pack(side="left", expand=True, fill="x", padx=(0, 12))
        ttk.Label(start_wrap, text="START DATE", style="Sub.TLabel").pack(
            anchor="w", pady=(0, 8)
        )
        start_picker = date_entry_cls(
            start_wrap, background=COLORS.accent, foreground="white", borderwidth=0
        )
        start_picker.pack(fill="x")
        self.start_date_picker = start_picker

        try:
            start_picker.bind("<<DateEntrySelected>>", self._on_start_date_selected)
        except Exception:
            pass

        # End Date
        end_wrap = ttk.Frame(range_row)
        end_wrap.pack(side="left", expand=True, fill="x", padx=(12, 0))
        ttk.Label(end_wrap, text="END DATE", style="Sub.TLabel").pack(
            anchor="w", pady=(0, 8)
        )
        end_picker = date_entry_cls(
            end_wrap, background=COLORS.accent, foreground="white", borderwidth=0
        )
        end_picker.pack(fill="x")
        self.end_date_picker = end_picker

    def _on_start_date_selected(self, event: Optional[object] = None) -> None:
        """Default end date to start date when needed."""

        if not self.start_date_picker or not self.end_date_picker:
            return

        try:
            start_dt = self.start_date_picker.get_date()
            end_dt = self.end_date_picker.get_date()
            if end_dt < start_dt:
                self.end_date_picker.set_date(start_dt)
        except Exception:
            # DateEntry implementations vary; never block the UI on this helper.
            return

    def _create_printer_row(self, parent: Union[ttk.Frame, ttk.LabelFrame]) -> None:
        """Create the printer selection row."""
        output_row = ttk.Frame(parent)
        output_row.pack(fill="x")
        ttk.Label(output_row, text="TARGET PRINTER", style="Sub.TLabel").pack(
            anchor="w", pady=(0, 8)
        )

        if win32print is None:
            logger.error("win32print is not available; printer enumeration disabled")

        all_printers = self._enumerate_printers()
        self._cached_printers = all_printers
        logger.debug(f"Found {len(all_printers)} printers")

        self.printer_var = tk.StringVar(value=DEFAULT_PRINTER_LABEL)
        printer_row = ttk.Frame(output_row)
        printer_row.pack(fill="x")

        self.printer_dropdown = ttk.OptionMenu(
            printer_row, self.printer_var, DEFAULT_PRINTER_LABEL, *all_printers
        )
        self.printer_dropdown.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ttk.Button(
            printer_row, text="Refresh", width=10, command=self.refresh_printers
        ).pack(side="right")

        if not all_printers:
            msg = "No printers found. Check connections."
            if win32print is None:
                msg = "Printing requires Windows with pywin32 installed (win32print unavailable)."
            ttk.Label(
                output_row, text=msg, style="Sub.TLabel", foreground=COLORS.error
            ).pack(anchor="w", pady=(4, 0))

    def _create_options_row(self, parent: Union[ttk.Frame, ttk.LabelFrame]) -> None:
        """Create advanced options row."""

        options_row = ttk.Frame(parent)
        options_row.pack(fill="x", pady=(16, 0))
        ttk.Label(options_row, text="OPTIONS", style="Sub.TLabel").pack(
            anchor="w", pady=(0, 6)
        )

        self.headers_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_row,
            text="Replace dates in headers/footers only (safer)",
            variable=self.headers_only_var,
        ).pack(anchor="w")

    def _create_footer(self, parent: ttk.Frame) -> None:
        """Create the action footer."""
        footer = ttk.Frame(parent)
        footer.pack(fill="x", side="bottom")

        # Status Label
        status_wrap = ttk.Frame(footer)
        status_wrap.pack(fill="x", pady=(0, 12))
        self.status_label = ttk.Label(
            status_wrap, text="System Ready", style="Sub.TLabel"
        )
        self.status_label.pack(side="left")

        ttk.Button(
            status_wrap, text="Open Logs", width=10, command=self.open_logs_folder
        ).pack(side="right")

        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            footer,
            variable=self.progress_var,
            maximum=PROGRESS_MAX,
            style="Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", pady=(0, 24))

        # Print Button
        self.print_btn = tk.Button(
            footer,
            text="START EXECUTION",
            bg=COLORS.accent,
            fg="#FFFFFF",
            font=FONTS.button,
            relief="flat",
            pady=18,
            cursor="hand2",
            activebackground=COLORS.accent_hover,
        )
        self.print_btn.pack(fill="x")

    def _create_path_row(
        self, parent: Union[ttk.Frame, ttk.LabelFrame], label: str, default_val: str
    ) -> ttk.Entry:
        """
        Create a path input row with browse button.

        Args:
            parent: Parent widget
            label: Label text
            default_val: Default path value

        Returns:
            The entry widget
        """
        wrap = ttk.Frame(parent)
        wrap.pack(fill="x", pady=8)
        ttk.Label(wrap, text=label, style="Sub.TLabel").pack(anchor="w", pady=(0, 4))

        row = ttk.Frame(wrap)
        row.pack(fill="x")

        entry = ttk.Entry(row)
        entry.insert(0, default_val)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ttk.Button(
            row, text="Browse", width=10, command=lambda: self._browse_folder(entry)
        ).pack(side="right")

        return entry

    def _browse_folder(self, entry: ttk.Entry) -> None:
        """
        Open folder browser dialog.

        Args:
            entry: Entry widget to update with selected path
        """
        path = filedialog.askdirectory()
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            logger.debug(f"Selected folder: {path}")

    def get_day_folder(self) -> str:
        """Get the day folder path."""
        return self.day_entry.get() if self.day_entry else ""

    def get_night_folder(self) -> str:
        """Get the night folder path."""
        return self.night_entry.get() if self.night_entry else ""

    def get_printer_name(self) -> str:
        """Get the selected printer name."""
        return self.printer_var.get() if self.printer_var else ""

    def get_available_printers(self) -> list[str]:
        """Return the available printers list (best-effort)."""
        return list(self._cached_printers)

    def get_headers_footers_only(self) -> bool:
        """Return True if date replacement should only touch headers/footers."""
        return bool(self.headers_only_var.get()) if self.headers_only_var else False

    def get_start_date(self) -> Optional[date]:
        """Get the start date."""
        return self.start_date_picker.get_date() if self.start_date_picker else None

    def get_end_date(self) -> Optional[date]:
        """Get the end date."""
        return self.end_date_picker.get_date() if self.end_date_picker else None

    def set_start_command(self, command: Callable[[], None]) -> None:
        """
        Set the command for the print button.

        Args:
            command: Function to call when button is clicked
        """
        if self.print_btn:
            self.print_btn.config(command=command)

    def set_print_button_state(self, state: Literal["normal", "disabled"]) -> None:
        """
        Set the print button state.

        Args:
            state: Either "normal" or "disabled"
        """
        if self.print_btn:
            self.print_btn.config(state=state)

    def update_status(self, message: str, progress: float) -> None:
        """
        Update the status label and progress bar.

        Args:
            message: Status message to display
            progress: Progress value (0-100)
        """
        if self.status_label:
            self.status_label.config(text=message)
        if self.progress_var:
            self.progress_var.set(progress)

    def show_error(self, title: str, message: str) -> None:
        """Show an error message box."""
        logger.error(f"{title}: {message}")
        messagebox.showerror(title, message)

    def show_warning(self, title: str, message: str) -> None:
        """Show a warning message box."""
        logger.warning(f"{title}: {message}")
        messagebox.showwarning(title, message)

    def show_info(self, title: str, message: str) -> None:
        """Show an info message box."""
        logger.info(f"{title}: {message}")
        messagebox.showinfo(title, message)

    def ask_yes_no(self, title: str, message: str) -> bool:
        """Ask the user a yes/no question."""
        return bool(messagebox.askyesno(title, message))

    def open_logs_folder(self) -> None:
        """Open the app data/log directory in the OS file explorer."""

        path = get_data_dir()
        try:
            path.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        try:
            if hasattr(os, "startfile"):
                os.startfile(str(path))  # type: ignore[attr-defined]
                return
        except Exception as e:
            logger.debug(f"os.startfile failed: {e}")

        try:
            if sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as e:
            logger.error(f"Could not open logs folder: {e}")

    def run(self) -> None:
        """Start the main UI loop."""
        logger.info("Starting UI main loop")
        self.root.mainloop()
