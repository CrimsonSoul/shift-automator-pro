import os
import json
import datetime
import time
import threading
import calendar
from tkinter import messagebox, filedialog, ttk
import tkinter as tk
from tkcalendar import DateEntry
import win32com.client
import win32print
import pythoncom

# Modern Aesthetics
COLORS = {
    "background": "#0f111a",  # Deep midnight
    "surface": "#1a1c27",     # Elevated surface
    "accent": "#00d2ff",      # Cyan glow
    "text_main": "#e6edf3",   # Soft white
    "text_dim": "#9198a1",    # Muted grey
    "success": "#23d160",     # Vibrant green
    "border": "#30363d",      # Subtle divider
    "secondary": "#21262d"    # Button hover
}

DEFAULT_CONFIG = {
    "day_folder": "",
    "night_folder": "",
    "printer_name": ""
}

CONFIG_FILE = "config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except:
            return DEFAULT_CONFIG
    return DEFAULT_CONFIG

def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=4)

def is_third_thursday(dt):
    """Returns True if the given date is the third Thursday of its month."""
    if dt.weekday() != 3: # 3 is Thursday
        return False
    # Get all thursdays in the month
    month_calendar = calendar.monthcalendar(dt.year, dt.month)
    thursdays = [week[3] for week in month_calendar if week[3] != 0]
    return dt.day == thursdays[2] if len(thursdays) >= 3 else False

class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shift Automator Pro")
        self.root.geometry("640x680")
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS["background"])
        self.config = load_config()
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        self.create_widgets()

    def configure_styles(self):
        # Global Styles
        self.style.configure("TFrame", background=COLORS["background"])
        self.style.configure("TLabel", background=COLORS["background"], foreground=COLORS["text_main"], font=("Segoe UI", 10))
        
        # Label Frame (Card Look)
        self.style.configure("TLabelframe", background=COLORS["background"], foreground=COLORS["accent"], bordercolor=COLORS["border"], lightcolor=COLORS["border"])
        self.style.configure("TLabelframe.Label", background=COLORS["background"], foreground=COLORS["accent"], font=("Segoe UI", 10, "bold"))
        
        # Entry
        self.style.configure("TEntry", fieldbackground=COLORS["surface"], foreground=COLORS["text_main"], insertcolor=COLORS["text_main"], borderwidth=0)
        
        # Buttons
        self.style.configure("TButton", background=COLORS["surface"], foreground=COLORS["text_main"], borderwidth=0, font=("Segoe UI", 9, "bold"))
        self.style.map("TButton", background=[("active", COLORS["secondary"])])
        
        # Progress Bar
        self.style.configure("Horizontal.TProgressbar", thickness=6, troughcolor="#21262d", background=COLORS["accent"])

        # Specialized Labels
        self.style.configure("Header.TLabel", font=("Segoe UI", 22, "bold"), foreground=COLORS["text_main"], background=COLORS["background"])
        self.style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground=COLORS["text_dim"], background=COLORS["background"])

    def create_widgets(self):
        container = ttk.Frame(self.root, padding="40")
        container.pack(fill="both", expand=True)

        # Branding
        header_frame = ttk.Frame(container)
        header_frame.pack(fill="x", pady=(0, 30))
        ttk.Label(header_frame, text="Shift Automator", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header_frame, text="Batch Schedule Management & Printing", style="Sub.TLabel").pack(anchor="w")

        # Folder Paths (Section 1)
        paths_frame = ttk.LabelFrame(container, text=" DIRECTORIES ", padding="20")
        paths_frame.pack(fill="x", pady=(0, 15))

        self.day_entry = self.create_path_row(paths_frame, "Day Shift", self.config["day_folder"])
        self.night_entry = self.create_path_row(paths_frame, "Night Shift", self.config["night_folder"])

        # Date Range (Section 2)
        dates_frame = ttk.LabelFrame(container, text=" DATE RANGE ", padding="20")
        dates_frame.pack(fill="x", pady=(0, 15))

        range_container = ttk.Frame(dates_frame)
        range_container.pack(fill="x")
        
        # Start Date
        start_col = ttk.Frame(range_container)
        start_col.pack(side="left", expand=True, fill="x", padx=(0, 10))
        ttk.Label(start_col, text="From:", style="Sub.TLabel").pack(anchor="w")
        self.start_date_picker = DateEntry(start_col, width=12, background=COLORS["accent"], foreground='white', borderwidth=0)
        self.start_date_picker.pack(fill="x", pady=(5, 0))

        # End Date
        end_col = ttk.Frame(range_container)
        end_col.pack(side="left", expand=True, fill="x", padx=(10, 0))
        ttk.Label(end_col, text="To:", style="Sub.TLabel").pack(anchor="w")
        self.end_date_picker = DateEntry(end_col, width=12, background=COLORS["accent"], foreground='white', borderwidth=0)
        self.end_date_picker.pack(fill="x", pady=(5, 0))

        # Printer (Section 3)
        printer_frame = ttk.LabelFrame(container, text=" OUTPUT ", padding="20")
        printer_frame.pack(fill="x", pady=(0, 20))
        
        self.printer_var = tk.StringVar(value=self.config["printer_name"])
        local_printers = [p[2] for p in win32print.EnumPrinters(2)]
        network_printers = [p[2] for p in win32print.EnumPrinters(4)]
        all_printers = sorted(list(set(local_printers + network_printers)))
        
        self.printer_dropdown = ttk.OptionMenu(printer_frame, self.printer_var, self.config["printer_name"] or "Choose Printer", *all_printers)
        self.printer_dropdown.pack(fill="x")

        # Execution Area
        exec_frame = ttk.Frame(container)
        exec_frame.pack(fill="x", side="bottom", pady=(20, 0))

        self.status_label = ttk.Label(exec_frame, text="Ready for execution", style="Sub.TLabel")
        self.status_label.pack(side="top", pady=(0, 5))

        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(exec_frame, variable=self.progress_var, maximum=100, style="Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 25))

        self.print_btn = tk.Button(exec_frame, text="START PROCESSING", command=self.start_thread, 
                                 bg=COLORS["success"], fg=COLORS["background"], font=("Segoe UI Variable Display", 11, "bold"),
                                 relief="flat", pady=15, cursor="hand2", activebackground=COLORS["accent"])
        self.print_btn.pack(fill="x")

    def create_path_row(self, parent, label, default_val):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=5)
        ttk.Label(row, text=label, width=12).pack(side="left")
        
        entry_frame = ttk.Frame(row)
        entry_frame.pack(side="left", fill="x", expand=True, padx=10)
        
        entry = ttk.Entry(entry_frame)
        entry.insert(0, default_val)
        entry.pack(fill="x")
        
        ttk.Button(row, text="Browse", width=10, command=lambda: self.browse_folder(entry)).pack(side="right")
        return entry

    def browse_folder(self, entry):
        path = filedialog.askdirectory()
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def start_thread(self):
        self.print_btn.config(state="disabled")
        threading.Thread(target=self.process_batch, daemon=True).start()

    def process_batch(self):
        pythoncom.CoInitialize()
        start_date = self.start_date_picker.get_date()
        end_date = self.end_date_picker.get_date()
        
        if end_date < start_date:
            messagebox.showerror("Error", "End date cannot be before start date.")
            self.root.after(0, lambda: self.print_btn.config(state="normal"))
            return

        day_folder = self.day_entry.get()
        night_folder = self.night_entry.get()
        printer = self.printer_var.get()

        if not day_folder or not night_folder:
            messagebox.showwarning("Incomplete Configuration", "Please select both Day and Night shift folders before starting.")
            self.root.after(0, lambda: self.print_btn.config(state="normal"))
            return

        self.config.update({"day_folder": day_folder, "night_folder": night_folder, "printer_name": printer})
        save_config(self.config)

        total_days = (end_date - start_date).days + 1
        word = None
        
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            for i in range(total_days):
                current_date = start_date + datetime.timedelta(days=i)
                day_name = current_date.strftime("%A")
                
                # Update status
                display_date = current_date.strftime("%m/%d/%Y")
                self.update_status(f"Processing {day_name} {display_date}...", (i / total_days) * 100)

                # Day Shift Logic
                day_file_id = "THIRD Thursday" if is_third_thursday(current_date) else day_name
                self.print_doc(word, day_folder, day_file_id, current_date, printer)
                
                # Night Shift Logic
                self.print_doc(word, night_folder, f"{day_name} Night", current_date, printer)

            self.update_status("Complete!", 100)
            messagebox.showinfo("Success", "All selected schedules have been processed and sent to the printer.")
        except Exception as e:
            messagebox.showerror("Processing Error", f"An error occurred: {str(e)}")
        finally:
            if word:
                try: word.Quit()
                except: pass
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.print_btn.config(state="normal"))

    def update_status(self, msg, progress):
        self.root.after(0, lambda: self.status_label.config(text=msg))
        self.root.after(0, lambda: self.progress_var.set(progress))

    def print_doc(self, word_app, folder, day_filename_part, current_date, printer_name):
        target_file = None
        files = os.listdir(folder)
        
        # 1. Try exact match first (e.g. "Thursday.docx") to avoid subset matches
        for f in files:
            if f.lower() == f"{day_filename_part.lower()}.docx":
                target_file = os.path.join(folder, f)
                break
        
        # 2. Fallback to "contains" if no exact match (supports older naming styles)
        if not target_file:
            for f in files:
                if day_filename_part.lower() in f.lower() and f.endswith(".docx"):
                    target_file = os.path.join(folder, f)
                    break
        
        if not target_file: return 

        try:
            doc = word_app.Documents.Open(target_file, False, False)
            if doc.ProtectionType != -1:
                try: doc.Unprotect()
                except: pass

            self.replace_with_regex(word_app, doc, current_date)
            
            word_app.ActivePrinter = printer_name
            doc.PrintOut(Background=False)
            self.safe_com_call(doc.Close, 0)
        except Exception as e:
            print(f"Error printing {target_file}: {e}")

    def safe_com_call(self, func, *args, retries=5, delay=1):
        for i in range(retries):
            try: return func(*args)
            except Exception as e:
                if "rejected" in str(e).lower() and i < retries - 1:
                    time.sleep(delay)
                    continue
                raise e

    def replace_with_regex(self, word_app, doc, current_date):
        """Turbo-optimized replacement using high-performance wildcards."""
        new_day = current_date.strftime("%A")
        new_month = current_date.strftime("%B")
        new_day_num = str(int(current_date.strftime("%d")))
        new_year = current_date.strftime("%Y")

        # We use [A-Za-z]@ which means "one or more letters" (much faster than *)
        # Patterns match: "DayName, Month Day, Year" and "DayName Month Day, Year"
        
        # 1. Day Shift Style (With Comma): "Sunday, January 04, 2026"
        self.execute_replace(doc, "[A-Za-z]@, [A-Za-z]@ [0-9]{1,2}, [0-9]{4}", f"{new_day}, {new_month} {new_day_num}, {new_year}")
        
        # 2. Night Shift Style (No Comma): "Saturday January 03, 2026"
        self.execute_replace(doc, "[A-Za-z]@ [A-Za-z]@ [0-9]{1,2}, [0-9]{4}", f"{new_day} {new_month} {new_day_num}, {new_year}")
        
        # 3. Fallback/Standard Style: "January 04, 2026"
        self.execute_replace(doc, "[A-Za-z]@ [0-9]{1,2}, [0-9]{4}", f"{new_month} {new_day_num}, {new_year}")

    def execute_replace(self, doc, find_text, replace_text):
        try:
            for story in doc.StoryRanges:
                self._run_find_replace(story, find_text, replace_text)
                nxt = story.NextStoryRange
                while nxt:
                    self._run_find_replace(nxt, find_text, replace_text)
                    nxt = nxt.NextStoryRange
        except: pass

    def _run_find_replace(self, range_obj, find_text, replace_text):
        f = range_obj.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        f.Execute(find_text, False, False, True, False, False, True, 1, False, replace_text, 2)

if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()
