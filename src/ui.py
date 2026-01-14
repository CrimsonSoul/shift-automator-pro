"""
UI components for Shift Automator application.

This module contains all Tkinter UI components and styling.
"""

import os
import tkinter as tk
from datetime import date, datetime
from tkinter import messagebox, filedialog, ttk, scrolledtext
from tkcalendar import DateEntry
from typing import Optional, Callable, List, Any, Union

# Platform-specific imports
try:
    import win32print
    HAS_WIN32PRINT = True
except ImportError:
    HAS_WIN32PRINT = False
    win32print = None  # type: ignore

from .constants import (
    COLORS, FONTS,
    WINDOW_WIDTH, WINDOW_HEIGHT, WINDOW_RESIZABLE,
    PROGRESS_MAX,
    PRINTER_ENUM_LOCAL, PRINTER_ENUM_NETWORK
)
from .logger import get_logger

logger = get_logger(__name__)

# Type aliases for callbacks
CommandCallback = Callable[[], None]
ConfigChangeCallback = Callable[[], None]
StatusUpdateCallback = Callable[[str, Optional[float]], None]


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
        self.style.theme_use('clam')
        self._configure_styles()

        # UI components
        self.day_entry: Optional[ttk.Entry] = None
        self.night_entry: Optional[ttk.Entry] = None
        self.start_date_picker: Optional[DateEntry] = None
        self.end_date_picker: Optional[DateEntry] = None
        self.printer_var: Optional[tk.StringVar] = None
        self.printer_dropdown: Optional[ttk.Combobox] = None
        self.log_widget: Optional[scrolledtext.ScrolledText] = None
        self.progress_var: Optional[tk.DoubleVar] = None
        self.progress: Optional[ttk.Progressbar] = None
        self.print_btn: Optional[tk.Button] = None
        self.cancel_btn: Optional[tk.Button] = None

        # Callback for configuration changes
        self._on_config_change: Optional[ConfigChangeCallback] = None

        # Create widgets
        self._create_widgets()

        logger.info("UI initialized")

    def _configure_styles(self) -> None:
        """Configure ttk styles for the application."""
        # Main background and surfaces
        self.style.configure("TFrame", background=COLORS.background)
        self.style.configure("Surface.TFrame", background=COLORS.surface)
        
        # Labels
        self.style.configure("TLabel", background=COLORS.background,
                           foreground=COLORS.text_main, font=FONTS.main)
        self.style.configure("Header.TLabel", font=FONTS.header,
                           foreground=COLORS.text_main, background=COLORS.background)
        self.style.configure("Sub.TLabel", font=FONTS.sub,
                           foreground=COLORS.text_dim, background=COLORS.background)
        self.style.configure("Surface.TLabel", background=COLORS.surface,
                           foreground=COLORS.text_main, font=FONTS.main)
        self.style.configure("SurfaceSub.TLabel", background=COLORS.surface,
                           foreground=COLORS.text_dim, font=FONTS.sub)

        # Labelframe (Relay Cards)
        self.style.configure("TLabelframe", background=COLORS.surface,
                           foreground=COLORS.text_dim, bordercolor=COLORS.border,
                           borderwidth=1, relief="solid")
        self.style.configure("TLabelframe.Label", background=COLORS.background,
                           foreground=COLORS.text_tertiary, font=FONTS.sub)

        # Entry fields
        self.style.configure("TEntry", fieldbackground=COLORS.surface_elevated,
                           foreground=COLORS.text_main, insertcolor=COLORS.text_main,
                           bordercolor=COLORS.border, lightcolor=COLORS.border,
                           darkcolor=COLORS.border, borderwidth=1)

        # Buttons (Relay Tactile Style)
        self.style.configure("TButton", background=COLORS.secondary,
                           foreground=COLORS.text_main, borderwidth=1,
                           bordercolor=COLORS.border, font=FONTS.bold, padding=(12, 6))
        self.style.map("TButton", 
                      background=[("active", COLORS.surface_elevated)],
                      bordercolor=[("active", COLORS.text_tertiary)])

        # Combobox (Dropdowns)
        self.style.configure("TCombobox", fieldbackground=COLORS.surface_elevated,
                           background=COLORS.surface_elevated, foreground=COLORS.text_main,
                           bordercolor=COLORS.border, arrowcolor=COLORS.text_dim)
        self.style.map("TCombobox", 
                      fieldbackground=[("readonly", COLORS.surface_elevated)],
                      selectbackground=[("readonly", COLORS.accent)],
                      selectforeground=[("readonly", COLORS.text_main)])

        # Progress bar
        self.style.configure("Horizontal.TProgressbar", thickness=6,
                           troughcolor=COLORS.border, background=COLORS.accent,
                           borderwidth=0)

        # Configure the dropdown list colors (popdown)
        self.root.option_add("*TCombobox*Listbox.background", COLORS.surface_elevated)
        self.root.option_add("*TCombobox*Listbox.foreground", COLORS.text_main)
        self.root.option_add("*TCombobox*Listbox.selectBackground", COLORS.accent)
        self.root.option_add("*TCombobox*Listbox.selectForeground", COLORS.text_main)
        self.root.option_add("*TCombobox*Listbox.font", FONTS.main)

    def _create_widgets(self) -> None:
        """Create all UI widgets."""
        # Background Container
        bg_container = ttk.Frame(self.root, padding="32")
        bg_container.pack(fill="both", expand=True)

        # Header Section
        self._create_header(bg_container)

        # Config Card
        self._create_config_card(bg_container)

        # Control Card
        self._create_control_card(bg_container)

        # Action Footer
        self._create_footer(bg_container)

    def _create_header(self, parent: ttk.Frame) -> None:
        """Create the header section."""
        header_row = ttk.Frame(parent)
        header_row.pack(fill="x", pady=(0, 32))
        ttk.Label(header_row, text="Shift Automator",
                 style="Header.TLabel").pack(anchor="w")
        ttk.Label(header_row, text="High-performance batch scheduling & printing",
                 style="Sub.TLabel").pack(anchor="w", pady=(4, 0))

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

    def _create_date_range_row(self, parent: Union[ttk.Frame, ttk.LabelFrame]) -> None:
        """Create the date range selection row."""
        range_row = ttk.Frame(parent)
        range_row.pack(fill="x", pady=(0, 20))

        # Date Picker Style
        date_style = {
            'background': COLORS.accent,
            'foreground': COLORS.text_main,
            'headersbackground': COLORS.surface_elevated,
            'headersforeground': COLORS.text_dim,
            'selectbackground': COLORS.accent,
            'selectforeground': COLORS.text_main,
            'normalbackground': COLORS.surface,
            'normalforeground': COLORS.text_main,
            'weekendbackground': COLORS.surface,
            'weekendforeground': COLORS.error,
            'othermonthforeground': COLORS.text_tertiary,
            'othermonthbackground': COLORS.surface,
            'othermonthweforeground': COLORS.text_tertiary,
            'othermonthwebackground': COLORS.surface,
            'borderwidth': 0
        }

        # Start Date
        start_wrap = ttk.Frame(range_row)
        start_wrap.pack(side="left", expand=True, fill="x", padx=(0, 12))
        ttk.Label(start_wrap, text="START DATE", style="Sub.TLabel").pack(anchor="w", pady=(0, 8))
        self.start_date_picker = DateEntry(start_wrap, **date_style)
        self.start_date_picker.pack(fill="x")

        # End Date
        end_wrap = ttk.Frame(range_row)
        end_wrap.pack(side="left", expand=True, fill="x", padx=(12, 0))
        ttk.Label(end_wrap, text="END DATE", style="Sub.TLabel").pack(anchor="w", pady=(0, 8))
        self.end_date_picker = DateEntry(end_wrap, **date_style)
        self.end_date_picker.pack(fill="x")

    def _create_printer_row(self, parent: Union[ttk.Frame, ttk.LabelFrame]) -> None:
        """Create the printer selection row."""
        output_row = ttk.Frame(parent)
        output_row.pack(fill="x")
        ttk.Label(output_row, text="TARGET PRINTER", style="Sub.TLabel").pack(anchor="w", pady=(0, 8))

        # Check platform compatibility
        if not HAS_WIN32PRINT:
            logger.error("win32print not available - application requires Windows")
            ttk.Label(
                output_row,
                text="Error: This application requires Windows to access printers.",
                foreground=COLORS.error
            ).pack(fill="x", pady=4)
            self.printer_var = tk.StringVar(value="")
            self.printer_dropdown = ttk.Combobox(
                output_row, textvariable=self.printer_var, values=["Not Available"], state="disabled"
            )
            self.printer_dropdown.pack(fill="x")
            return

        # Get available printers
        try:
            local_printers = [p[2] for p in win32print.EnumPrinters(PRINTER_ENUM_LOCAL)]  # type: ignore
            network_printers = [p[2] for p in win32print.EnumPrinters(PRINTER_ENUM_NETWORK)]  # type: ignore
            all_printers = sorted(list(set(local_printers + network_printers)))
            logger.debug(f"Found {len(all_printers)} printers")
        except Exception as e:
            logger.error(f"Error enumerating printers: {e}")
            all_printers = []

        self.printer_var = tk.StringVar(value="")
        # Add trace to save config when printer selection changes
        self.printer_var.trace_add("write", lambda *args: self._on_printer_change())
        
        self.printer_dropdown = ttk.Combobox(
            output_row, 
            textvariable=self.printer_var, 
            values=all_printers,
            state="readonly"
        )
        self.printer_dropdown.pack(fill="x")

    def _create_footer(self, parent: ttk.Frame) -> None:
        """Create the action footer."""
        footer = ttk.Frame(parent)
        footer.pack(fill="x", side="bottom")

        # Log Section
        log_wrap = ttk.Frame(footer)
        log_wrap.pack(fill="x", pady=(0, 12))
        ttk.Label(log_wrap, text="ACTIVITY LOG", style="Sub.TLabel").pack(anchor="w", pady=(0, 4))
        
        self.log_widget = scrolledtext.ScrolledText(
            log_wrap, height=6, font=("Consolas", 9),
            bg=COLORS.surface_elevated, fg=COLORS.text_dim,
            insertbackground=COLORS.text_main, borderwidth=0,
            highlightthickness=1, highlightbackground=COLORS.border,
            padx=10, pady=10
        )
        self.log_widget.pack(fill="x")
        self.log_widget.config(state="disabled")

        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(footer, variable=self.progress_var,
                                       maximum=PROGRESS_MAX, style="Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(12, 24))

        # Button Row
        button_row = ttk.Frame(footer)
        button_row.pack(fill="x")

        # Print Button
        self.print_btn = tk.Button(
            button_row, text="START EXECUTION",
            bg=COLORS.accent, fg="#FFFFFF", font=FONTS.button,
            relief="flat", pady=18, cursor="hand2", activebackground=COLORS.accent_hover
        )
        self.print_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        # Cancel Button
        self.cancel_btn = tk.Button(
            button_row, text="CANCEL",
            bg=COLORS.secondary, fg=COLORS.text_main, font=FONTS.button,
            relief="flat", pady=18, cursor="hand2", activebackground=COLORS.surface_elevated,
            state="disabled"
        )
        self.cancel_btn.pack(side="right", fill="x", expand=True, padx=(6, 0))

    def _create_path_row(self, parent: Union[ttk.Frame, ttk.LabelFrame], label: str, default_val: str) -> ttk.Entry:
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

        ttk.Button(row, text="Browse", width=10,
                  command=lambda: self._browse_folder(entry)).pack(side="right")

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
            # Trigger config change callback if set
            if self._on_config_change:
                self._on_config_change()

    def _on_printer_change(self) -> None:
        """Handle printer selection change."""
        # Trigger config change callback if set
        if self._on_config_change:
            self._on_config_change()

    def get_day_folder(self) -> str:
        """Get the day folder path."""
        return self.day_entry.get() if self.day_entry else ""

    def get_night_folder(self) -> str:
        """Get the night folder path."""
        return self.night_entry.get() if self.night_entry else ""

    def get_printer_name(self) -> str:
        """Get the selected printer name."""
        return self.printer_var.get() if self.printer_var else ""

    def get_start_date(self) -> Optional[date]:
        """
        Get the start date with error handling.

        Returns:
            The start date from the date picker, or None if not available
        """
        if not self.start_date_picker:
            return None
        try:
            return self.start_date_picker.get_date()
        except Exception as e:
            logger.warning(f"Error getting start date: {e}")
            return None

    def get_end_date(self) -> Optional[date]:
        """
        Get the end date with error handling.

        Returns:
            The end date from the date picker, or None if not available
        """
        if not self.end_date_picker:
            return None
        try:
            return self.end_date_picker.get_date()
        except Exception as e:
            logger.warning(f"Error getting end date: {e}")
            return None

    def set_start_command(self, command: CommandCallback) -> None:
        """
        Set the command for the print button.

        Args:
            command: Function to call when button is clicked
        """
        if self.print_btn:
            self.print_btn.config(command=command)

    def set_cancel_command(self, command: CommandCallback) -> None:
        """
        Set the command for the cancel button.

        Args:
            command: Function to call when button is clicked
        """
        if self.cancel_btn:
            self.cancel_btn.config(command=command)

    def set_config_change_callback(self, callback: ConfigChangeCallback) -> None:
        """
        Set a callback to be called when configuration changes.

        Args:
            callback: Function to call when configuration changes
        """
        self._on_config_change = callback

    def set_print_button_state(self, state: Any) -> None:
        """
        Set the print button state.

        Args:
            state: Either "normal" or "disabled"
        """
        if self.print_btn:
            self.print_btn.config(state=state)

    def set_cancel_button_state(self, state: Any) -> None:
        """
        Set the cancel button state.

        Args:
            state: Either "normal" or "disabled"
        """
        if self.cancel_btn:
            self.cancel_btn.config(state=state)

    def update_progress(self, progress: float) -> None:
        """
        Update the progress bar.

        Args:
            progress: Progress value (0-100)
        """
        if self.progress_var is not None:
            self.progress_var.set(progress)

    def log(self, message: str) -> None:
        """
        Append a message to the log widget.

        Args:
            message: Message to log
        """
        if self.log_widget:
            self.log_widget.config(state="normal")
            self.log_widget.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
            self.log_widget.see(tk.END)
            self.log_widget.config(state="disabled")

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

    def run(self) -> None:
        """Start the main UI loop."""
        logger.info("Starting UI main loop")
        self.root.mainloop()
