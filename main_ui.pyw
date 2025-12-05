# main_ui.pyw
"""
LUX Dynamics Thermal Temp Controller Logger

- GUI for configuring tests, channels, Arduino ambient control.
- Logs TC-08 + Arduino data to CSV and colored Excel.
- Opens a separate live graph window (graph_window.LiveGraphWindow).
- Shows channel trends with direction, average, and delta, e.g.:
    CH1: stable (avg 54.2 °C, Δ=+0.7 °C)
"""
import sys
import subprocess
import time
import csv
import os
import math
import tkinter as tk
# ---------------- Dependency checking ---------------- #
# pip_name, import_name
DEPENDENCIES = [
    ("openpyxl", "openpyxl"),          # for colored Excel output
    ("pywin32", "win32com.client"),    # for desktop shortcut creation
    ("pyserial", "serial"),            # for Arduino ambient controller
]


def check_and_install_dependencies():
    """
    Check for important Python packages and optionally auto-install them.

    - If everything is present: do nothing.
    - If some are missing: ask the user if they want to install.
    - If install succeeds for openpyxl: re-import it so Excel coloring works.
    - If some installs fail: app still runs, those features are just limited.
    """
    missing: list[tuple[str, str]] = []

    for pip_name, import_name in DEPENDENCIES:
        try:
            __import__(import_name.split(".")[0])
        except ImportError:
            missing.append((pip_name, import_name))

    if not missing:
        return  # all good

    # Build message text
    lines = [
        "This app is missing some Python packages needed for certain features:",
        "",
    ]
    for pip_name, import_name in missing:
        lines.append(f"• {pip_name}  (import '{import_name}')")
    lines.append("")
    lines.append("Would you like to try installing them automatically now?")

    msg = "\n".join(lines)

    if not messagebox.askyesno("Missing Python packages", msg):
        # User declined auto-install; keep going with reduced functionality.
        return

    python_exe = sys.executable or "python"
    failed: list[str] = []

    for pip_name, import_name in missing:
        try:
            subprocess.check_call([python_exe, "-m", "pip", "install", pip_name])
        except Exception:
            failed.append(pip_name)

    # Try to re-import openpyxl so HAVE_OPENPYXL becomes True if it was just installed.
    global HAVE_OPENPYXL, Workbook, PatternFill, Border, Side
    try:
        from openpyxl import Workbook as WB
        from openpyxl.styles import PatternFill as PF, Border as BD, Side as SD
        Workbook = WB
        PatternFill = PF
        Border = BD
        Side = SD
        HAVE_OPENPYXL = True
    except Exception:
        pass

    if failed:
        messagebox.showwarning(
            "Some packages not installed",
            "The app tried to install these packages but failed:\n\n"
            + "\n".join(f"• {name}" for name in failed)
            + "\n\nYou can still use most functionality, but some features may be disabled."
        )

from tkinter import ttk, messagebox

from logger_core import (
    TC08Interface,
    ArduinoInterface,
    TREND_WINDOW_DEFAULT,
    TREND_THRESHOLD_DEFAULT,
    SAMPLE_INTERVAL,
)
from graph_window import LiveGraphWindow

# Excel support...

# ---------------- App / shortcut constants ---------------- #
APP_NAME = "LUX Thermal Logger"          # Window title / shortcut name
SHORTCUT_NAME = "LUX Thermal Logger.lnk" # Name of the .lnk file on Desktop
ICON_FILENAME = "lux_logo.ico"           # Icon file sitting next to main_ui.pyw


def _get_windows_desktop() -> str:
    """
    Return the user's Desktop folder path on Windows, using the shell API
    if available, falling back to %USERPROFILE%\\Desktop.
    """
    if os.name != "nt":
        return os.path.join(os.path.expanduser("~"), "Desktop")

    try:
        import ctypes
        from ctypes import wintypes

        SHGFP_TYPE_CURRENT = 0
        CSIDL_DESKTOPDIRECTORY = 0x10

        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(
            None, CSIDL_DESKTOPDIRECTORY, None, SHGFP_TYPE_CURRENT, buf
        )
        return buf.value
    except Exception:
        # Fallback if anything weird happens
        return os.path.join(os.path.expanduser("~"), "Desktop")


def ensure_desktop_shortcut():
    """
    Create a Desktop shortcut to this script with your logo as the icon.

    - Only runs on Windows.
    - Only creates it if it doesn't already exist.
    - Uses lux_logo.ico in the same folder as main_ui.pyw if found.
    """
    if os.name != "nt":
        return

    try:
        from win32com.client import Dispatch
    except ImportError:
        # pywin32 not installed -> silently skip
        return

    desktop = _get_windows_desktop()
    shortcut_path = os.path.join(desktop, SHORTCUT_NAME)

    # If it already exists, don't recreate
    if os.path.exists(shortcut_path):
        return

    # Target = EXE when frozen, otherwise this .pyw file
    if getattr(sys, "frozen", False):
        target = sys.executable
        icon_location = f"{target},0"   # use EXE's embedded icon
    else:
        target = os.path.abspath(__file__)
        workdir = os.path.dirname(target)
        icon_path = os.path.join(workdir, ICON_FILENAME)
        if os.path.exists(icon_path):
            icon_location = f"{icon_path},0"
        else:
            icon_location = f"{target},0"

    workdir = os.path.dirname(target)

    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = workdir
    shortcut.IconLocation = icon_location
    shortcut.save()

from datetime import datetime
from typing import Dict, List, Tuple

import tkinter as tk
from tkinter import ttk, messagebox

from logger_core import (
    TC08Interface,
    ArduinoInterface,
    TREND_WINDOW_DEFAULT,
    TREND_THRESHOLD_DEFAULT,
    SAMPLE_INTERVAL,
)
from graph_window import LiveGraphWindow

# Excel support (for pretty colored columns)
try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Border, Side
    HAVE_OPENPYXL = True
except ImportError:
    Workbook = None
    PatternFill = None
    Border = None
    Side = None
    HAVE_OPENPYXL = False


# ---------------- File / output constants ---------------- #

OUTPUT_FOLDER = r"Z:\ENGINEERING\Product Development\Thermal Testing 2025"
"""Preferred root folder for logs; falls back to ./logs if Z: is unavailable."""


# ---------------- Helper functions ---------------- #

def get_unique_csv_path(folder: str, base_name: str) -> str:
    """
    Return a unique CSV path in 'folder' based on 'base_name'.

    Example:
        base_name = '2025-11-21 Thermal Test'
          -> '2025-11-21 Thermal Test.csv' (if free)
          -> '2025-11-21 Thermal Test_1.csv'
          -> '2025-11-21 Thermal Test_2.csv', etc.
    """
    path = os.path.join(folder, base_name + ".csv")
    if not os.path.exists(path):
        return path
    i = 1
    while True:
        alt = os.path.join(folder, f"{base_name}_{i}.csv")
        if not os.path.exists(alt):
            return alt
        i += 1


def resolve_output_folder() -> str:
    """
    Use Z: folder if available; otherwise local ./logs.
    """
    if os.path.isdir(OUTPUT_FOLDER):
        return OUTPUT_FOLDER
    fallback = os.path.join(os.getcwd(), "logs")
    os.makedirs(fallback, exist_ok=True)
    return fallback


def apply_column_colors(ws):
    """
    Color each column from the header row downward with a unique pale solid color
    and add bolder grid lines so columns stand out.
    """
    if not HAVE_OPENPYXL or PatternFill is None:
        return

    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            if cell.value == "timestamp":
                header_row_idx = cell.row
                break
        if header_row_idx is not None:
            break
    if header_row_idx is None:
        return

    header_cells = list(ws[header_row_idx])
    num_cols = len(header_cells)
    if num_cols == 0:
        return

    palette = [
        "FFCCCC", "FFE5CC", "FFF2CC", "E5FFCC",
        "CCFFFF", "CCE5FF", "E5CCFF", "FFCCF2", "E6E6FA",
    ]

    bold_border = Border(
        left=Side(style="medium", color="000000"),
        right=Side(style="medium", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for col_idx, cell in enumerate(header_cells, start=1):
        color_hex = palette[(col_idx - 1) % len(palette)]
        fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

        for row_idx in range(header_row_idx, ws.max_row + 1):
            c = ws.cell(row=row_idx, column=col_idx)
            c.fill = fill
            c.border = bold_border


def create_colored_excel(csv_path: str):
    """
    Read the CSV and create a colored Excel file (.xlsx) with each column
    tinted with a solid pale color and bold column borders.
    """
    if not HAVE_OPENPYXL or Workbook is None:
        print("openpyxl not available → skipping colored Excel export.")
        return

    xlsx_path = os.path.splitext(csv_path)[0] + ".xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "TC08 Log"

    with open(csv_path, newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            ws.append(row)

    apply_column_colors(ws)
    wb.save(xlsx_path)
    print(f"Colored Excel copy saved as:\n  {xlsx_path}")


def fmt_val(val):
    """
    Format numeric value to 2 decimal places for CSV/Excel.
    Returns '' for NaN / None so it doesn't blow up.
    """
    try:
        if val is None:
            return ""
        if isinstance(val, float) and math.isnan(val):
            return ""
        return f"{float(val):.2f}"
    except (TypeError, ValueError):
        return ""


# ---------------- Main GUI App ---------------- #

class ThermalLoggerApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("LUX Thermal Logger")
        self.geometry("920x620")

        # Hardware / file state
        self.logger = None
        self.csv_file = None
        self.csv_writer = None
        self.arduino = None

        # Run state
        self.is_logging = False
        self.start_time = None
        self.duration_seconds = None
        self.data_filename = None
        self.active_channels: List[Tuple[int, str]] = []
        self.use_arduino_flag = False
        self.ambient_setpoint_value = None

        # Live graph window
        self.graph_window: LiveGraphWindow | None = None

        # Trend detection state
        self.channel_history: Dict[int, List[float]] = {}
        self.trend_window = TREND_WINDOW_DEFAULT
        self.trend_threshold = TREND_THRESHOLD_DEFAULT

        # Handles
        self.status_label = None  # for red/black error state
        self.summary_header_text = ""

        # Build UI
        self._build_vars()
        self._build_ui()
        self.set_status("Idle.")
        self.after(0, self._post_init)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- Tkinter variable setup ---------- #
    def _post_init(self):
        """
        Run once after the Tk window is created:
        - Check/install Python dependencies.
        - Create Desktop shortcut (if pywin32 is available).
        """
        try:
            check_and_install_dependencies()
        except Exception:
            # Don't crash the app because dependency check exploded
            pass

        try:
            ensure_desktop_shortcut()
        except Exception:
            # Shortcut creation is non-fatal
            pass
# ---------- Tkinter variable setup ---------- #
    def _build_vars(self):
        # Metadata
        self.test_name_var = tk.StringVar()
        self.tester_var = tk.StringVar()
        self.fixture_var = tk.StringVar()
        self.notes_var = tk.StringVar()

        # Channels
        self.include_cj_var = tk.BooleanVar(value=False)
        self.num_inputs_var = tk.IntVar(value=2)
        self.ch_name_vars = [tk.StringVar(value=f"CH{i}") for i in range(1, 9)]

        # Arduino
        self.use_arduino_var = tk.BooleanVar(value=False)
        self.arduino_port_var = tk.StringVar(value="COM5")
        self.ambient_setpoint_var = tk.StringVar(value="25")

        # File / run settings
        today_str = datetime.now().strftime("%Y-%m-%d")
        default_name = f"{today_str} Thermal Test"
        self.base_name_var = tk.StringVar(value=default_name)
        self.duration_minutes_var = tk.StringVar(value="")

        # Status / summary
        self.status_var = tk.StringVar(value="Idle.")
        self.last_line_var = tk.StringVar(value="No data yet.")
        self.summary_var = tk.StringVar(value="No configuration yet.")

        # Channel trend text
        self.channel_trends_var = tk.StringVar(
            value="Channel temperature trends will appear here once data arrives."
        )

        # Trend settings UI-controlled
        self.trend_window_var = tk.StringVar(value=str(self.trend_window))
        self.trend_threshold_var = tk.StringVar(value=f"{self.trend_threshold:.1f}")

        # Output naming / path
        self.append_datetime_var = tk.BooleanVar(value=False)
        self.output_path_var = tk.StringVar(value="")

    # ---------- Status helper ---------- #

    def set_status(self, text: str, is_error: bool = False):
        self.status_var.set(text)
        if self.status_label is not None:
            self.status_label.configure(foreground=("red" if is_error else "black"))

    # ---------- UI layout ---------- #

    def _build_ui(self):
        # Top title bar
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(
            top,
            text="Thermal Temp Controller Logger",
            font=("Century Gothic", 16, "bold")
        ).pack(side="left")
        right_info = ttk.Frame(top)
        right_info.pack(side="right", anchor="e")
        ttk.Label(
            right_info, text="LUX Dynamics",
            font=("Century Gothic", 12, "bold")
        ).pack(anchor="e")
        ttk.Label(
            right_info, text="Kailani Puava Alarcon",
            font=("Century Gothic", 10)
        ).pack(anchor="e")

        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # Left column: metadata + channels + trends
        left = ttk.Frame(main)
        left.pack(side="left", fill="y", padx=(0, 12))

        meta = ttk.LabelFrame(left, text="Test Metadata", padding=10)
        meta.pack(fill="x", pady=(0, 10))

        ttk.Label(meta, text="Test name:").grid(row=0, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.test_name_var, width=28).grid(row=0, column=1, sticky="w")

        ttk.Label(meta, text="Tester:").grid(row=1, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.tester_var, width=28).grid(row=1, column=1, sticky="w")

        ttk.Label(meta, text="Fixture:").grid(row=2, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.fixture_var, width=28).grid(row=2, column=1, sticky="w")

        ttk.Label(meta, text="Notes:").grid(row=3, column=0, sticky="ne")
        ttk.Entry(meta, textvariable=self.notes_var, width=28).grid(row=3, column=1, sticky="w")

        ch_frame = ttk.LabelFrame(left, text="TC-08 Channels + Trends", padding=10)
        ch_frame.pack(fill="x")

        ttk.Checkbutton(
            ch_frame,
            text="Include internal sensor (channel 0 / CJ)",
            variable=self.include_cj_var
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(ch_frame, text="# of inputs to log (1–8):").grid(row=1, column=0, sticky="e")
        ttk.Spinbox(
            ch_frame, from_=0, to=8, textvariable=self.num_inputs_var,
            width=5
        ).grid(row=1, column=1, sticky="w")

        row = 2
        for i in range(1, 9):
            ttk.Label(ch_frame, text=f"Input {i} name:").grid(row=row, column=0, sticky="e")
            ttk.Entry(ch_frame, textvariable=self.ch_name_vars[i - 1], width=20).grid(
                row=row, column=1, sticky="w"
            )
            row += 1

        # Trend settings
        ttk.Label(ch_frame, text="Trend window (samples):").grid(row=row, column=0, sticky="e")
        ttk.Entry(ch_frame, textvariable=self.trend_window_var, width=8).grid(
            row=row, column=1, sticky="w"
        )
        row += 1

        ttk.Label(ch_frame, text="Stable band (°C):").grid(row=row, column=0, sticky="e")
        ttk.Entry(ch_frame, textvariable=self.trend_threshold_var, width=8).grid(
            row=row, column=1, sticky="w"
        )
        row += 1

        # Trends text
        ttk.Label(
            ch_frame,
            textvariable=self.channel_trends_var,
            justify="left",
            foreground="gray"
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Right column: Arduino + run settings + summary + status
        right = ttk.Frame(main)
        right.pack(side="left", fill="both", expand=True)

        ar_frame = ttk.LabelFrame(right, text="Arduino Ambient Control", padding=10)
        ar_frame.pack(fill="x")

        ttk.Checkbutton(
            ar_frame,
            text="Use Arduino for ambient control/logging",
            variable=self.use_arduino_var
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(ar_frame, text="COM port (e.g. COM5 or 5):").grid(row=1, column=0, sticky="e")
        ttk.Entry(ar_frame, textvariable=self.arduino_port_var, width=12).grid(
            row=1, column=1, sticky="w"
        )

        ttk.Label(ar_frame, text="Ambient setpoint (°C):").grid(row=2, column=0, sticky="e")
        ttk.Entry(ar_frame, textvariable=self.ambient_setpoint_var, width=12).grid(
            row=2, column=1, sticky="w"
        )

        run_frame = ttk.LabelFrame(right, text="Run Settings", padding=10)
        run_frame.pack(fill="x", pady=(10, 0))

        ttk.Label(run_frame, text="Output folder:").grid(row=0, column=0, sticky="ne")
        self.output_folder_label = ttk.Label(
            run_frame,
            text=resolve_output_folder(),
            wraplength=360,
            justify="left"
        )
        self.output_folder_label.grid(row=0, column=1, sticky="w")

        ttk.Label(run_frame, text="Base file name:").grid(row=1, column=0, sticky="e")
        ttk.Entry(run_frame, textvariable=self.base_name_var, width=32).grid(
            row=1, column=1, sticky="w"
        )

        self.append_datetime_check = ttk.Checkbutton(
            run_frame,
            text="Append start time to file name",
            variable=self.append_datetime_var
        )
        self.append_datetime_check.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 4))

        ttk.Label(run_frame, text="Duration (minutes, blank = unlimited):").grid(
            row=3, column=0, sticky="e"
        )
        ttk.Entry(run_frame, textvariable=self.duration_minutes_var, width=12).grid(
            row=3, column=1, sticky="w"
        )

        btn_frame = ttk.Frame(run_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0))

        self.start_button = ttk.Button(btn_frame, text="Start Logging", command=self.start_logging)
        self.start_button.pack(side="left", padx=(0, 10))

        self.stop_button = ttk.Button(btn_frame, text="Stop Logging", command=self.on_stop)
        self.stop_button.pack(side="left")
        self.stop_button["state"] = "disabled"

        open_graph_btn = ttk.Button(
            run_frame,
            text="Open Live Graph Window",
            command=self.ensure_graph_window
        )
        open_graph_btn.grid(row=5, column=0, columnspan=2, pady=(10, 0))

        ttk.Label(run_frame, text="Full output path:").grid(row=6, column=0, sticky="ne", pady=(8, 0))
        self.output_path_entry = ttk.Entry(run_frame, textvariable=self.output_path_var, width=42)
        self.output_path_entry.grid(row=6, column=1, sticky="w", pady=(8, 0))
        self.output_path_entry.configure(state="readonly")

        summary_frame = ttk.LabelFrame(right, text="Current Configuration", padding=10)
        summary_frame.pack(fill="both", expand=True, pady=(10, 0))
        ttk.Label(
            summary_frame,
            textvariable=self.summary_var,
            justify="left",
            wraplength=420
        ).pack(anchor="w")

        # Status section at the bottom
        status_frame = ttk.LabelFrame(self, text="Status", padding=10)
        status_frame.pack(fill="x", side="bottom")

        self.status_label = ttk.Label(status_frame, textvariable=self.status_var)
        self.status_label.pack(anchor="w")
        ttk.Label(status_frame, text="Last reading:").pack(anchor="w")
        ttk.Label(
            status_frame,
            textvariable=self.last_line_var,
            wraplength=860
        ).pack(anchor="w")

    # ---------- Graph window ---------- #

    def ensure_graph_window(self):
        if self.graph_window is None or not self.graph_window.winfo_exists():
            self.graph_window = LiveGraphWindow(self)
            if self.active_channels:
                self.graph_window.set_channels(self.active_channels)

    # ---------- Start logging ---------- #

    def start_logging(self):
        if self.is_logging:
            messagebox.showinfo("Logging", "Already logging.")
            return

        test_name = self.test_name_var.get().strip() or "Untitled Test"
        tester = self.tester_var.get().strip() or "Unknown"
        fixture = self.fixture_var.get().strip() or "N/A"
        notes = self.notes_var.get().strip()

        # Channels
        try:
            num_inputs = int(self.num_inputs_var.get())
        except ValueError:
            messagebox.showerror("Error", "Number of inputs must be a number between 0 and 8.")
            return
        if not (0 <= num_inputs <= 8):
            messagebox.showerror("Error", "Number of inputs must be between 0 and 8.")
            return

        channels: List[Tuple[int, str]] = []
        if self.include_cj_var.get():
            channels.append((0, "CJ"))

        for i in range(1, num_inputs + 1):
            name = self.ch_name_vars[i - 1].get().strip()
            if not name:
                name = f"CH{i}"
            channels.append((i, name))

        if not channels:
            messagebox.showerror("Error", "You must log at least one channel.")
            return

        self.active_channels = channels

        # Trend settings from UI
        try:
            tw_str = self.trend_window_var.get().strip()
            tw = int(tw_str)
            if tw < 2:
                raise ValueError
            self.trend_window = tw
        except Exception:
            messagebox.showerror(
                "Trend settings error",
                "Trend window (samples) must be an integer ≥ 2."
            )
            return

        try:
            band_str = self.trend_threshold_var.get().strip()
            band = float(band_str)
            if band <= 0:
                raise ValueError
            self.trend_threshold = band
        except Exception:
            messagebox.showerror(
                "Trend settings error",
                "Stable band (°C) must be a positive number."
            )
            return

        # Arduino
        self.use_arduino_flag = False
        self.ambient_setpoint_value = None
        if self.use_arduino_var.get():
            port_input = self.arduino_port_var.get().strip()
            if not port_input:
                messagebox.showerror("Arduino error", "Please enter a COM port (e.g. COM5 or 5).")
                return

            if port_input.upper().startswith("COM"):
                port_name = port_input.upper()
            else:
                port_name = f"COM{port_input}"

            sp_str = self.ambient_setpoint_var.get().strip()
            try:
                sp = float(sp_str)
            except ValueError:
                messagebox.showerror("Arduino error. Get Kailani.", "Ambient setpoint must be a number.")
                return

            try:
                self.arduino = ArduinoInterface(port_name)
                self.use_arduino_flag = True
                self.ambient_setpoint_value = sp
                self.arduino.set_hold(sp)
            except Exception as e:
                messagebox.showerror(
                    "Arduino error. Get Kailani.",
                    f"Failed to connect to Arduino on {port_name}:\n{e}"
                )
                self.arduino = None
                self.use_arduino_flag = False

        # Output folder & filename
        output_folder = resolve_output_folder()
        self.output_folder_label.config(text=output_folder)

        base_name = self.base_name_var.get().strip()
        if not base_name:
            today_str = datetime.now().strftime("%Y-%m-%d")
            base_name = f"{today_str} Thermal Test"

        if self.append_datetime_var.get():
            time_str = datetime.now().strftime("%H-%M-%S")
            base_name = f"{base_name} {time_str}"

        self.base_name_var.set(base_name)
        self.data_filename = get_unique_csv_path(output_folder, base_name)
        self.output_path_var.set(self.data_filename)

        # Duration
        duration_str = self.duration_minutes_var.get().strip()
        if duration_str == "":
            self.duration_seconds = None
        else:
            try:
                minutes = float(duration_str)
                if minutes <= 0:
                    raise ValueError
                self.duration_seconds = minutes * 60.0
            except ValueError:
                messagebox.showerror(
                    "Error",
                    "Duration must be a positive number of minutes or left blank."
                )
                return

        # Open TC-08
        try:
            self.logger = TC08Interface()
        except Exception as e:
            messagebox.showerror("TC-08 error. Get Kailani.", f"Could not open TC-08:\n{e}")
            self.logger = None
            self.set_status("TC-08 error: could not open device.", is_error=True)
            return

        # Open CSV
        try:
            self.csv_file = open(self.data_filename, mode="w", newline="")
            self.csv_writer = csv.writer(self.csv_file)
        except Exception as e:
            messagebox.showerror("File error. Get Kailani.", f"Could not open CSV file for writing:\n{e}")
            if self.logger is not None:
                self.logger.close()
            self.logger = None
            self.set_status("File error: could not open CSV for writing.", is_error=True)
            return

        # Write header
        meta_text = (
            f"Test: {test_name} | "
            f"Tester: {tester} | "
            f"Fixture: {fixture} | "
            f"Notes: {notes}"
        )
        if self.ambient_setpoint_value is not None:
            meta_text += f" | Ambient setpoint: {self.ambient_setpoint_value:.2f} °C"
        self.csv_writer.writerow([meta_text])
        self.csv_writer.writerow([])

        header = ["timestamp"]
        if self.use_arduino_flag:
            header.append("Arduino_Temp")
        for _, name in self.active_channels:
            header.append(f"{name}_C")
        self.csv_writer.writerow(header)
        self.csv_file.flush()

        summary_lines = [
            f"Output file: {os.path.basename(self.data_filename)}",
            f"Test: {test_name}",
            f"Tester: {tester}",
            f"Fixture: {fixture}",
            (
                f"Ambient setpoint: {self.ambient_setpoint_value:.2f} °C"
                if self.ambient_setpoint_value is not None
                else "Ambient setpoint: N/A"
            ),
            "Channels:",
        ]
        for ch, name in self.active_channels:
            summary_lines.append(f"  Input {ch}: {name}")

        self.summary_header_text = "\n".join(summary_lines)
        self.summary_var.set(self.summary_header_text)

        # Reset channel history each run
        self.channel_history = {}
        self.channel_trends_var.set(
            f"Channel temperature trends (last ~{self.trend_window} readings, "
            f"stable within ±{self.trend_threshold:.1f} °C) will appear here once data arrives."
        )

        # Set up graph window
        self.ensure_graph_window()
        if self.graph_window is not None and self.graph_window.winfo_exists():
            self.graph_window.set_channels(self.active_channels)

        # Start run
        self.start_time = time.time()
        self.is_logging = True
        self.set_status("Logging...")
        self.last_line_var.set("No data yet.")
        self.start_button["state"] = "disabled"
        self.stop_button["state"] = "normal"

        # First poll
        self.after(int(SAMPLE_INTERVAL * 1000), self.poll_once)

    # ---------- Poll loop ---------- #

    def poll_once(self):
        if not self.is_logging:
            return

        # Read TC-08 (non-fatal errors go to status bar)
        try:
            temps = self.logger.read() if self.logger is not None else {}
        except Exception as e:
            self.set_status(f"TC-08 read error: {e}", is_error=True)
            self.after(int(SAMPLE_INTERVAL * 1000), self.poll_once)
            return

        # Clear previous error if we recovered
        if self.status_var.get().startswith("TC-08 read error"):
            self.set_status("Logging...")

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row = [ts]
        display_vals: List[str] = []

        # Arduino
        if self.use_arduino_flag and self.arduino is not None:
            ar_temp, ar_hold, ar_pwm = self.arduino.poll()
            row.append(fmt_val(ar_temp))
            if ar_temp is not None:
                display_vals.append(
                    f"Arduino={ar_temp:.2f}°C (hold={ar_hold:.2f}°C, PWM={ar_pwm:.0f})"
                )
            else:
                display_vals.append("Arduino=NaN")

        # TC-08 channels
        for ch, name in self.active_channels:
            val = temps.get(ch, float("nan"))
            row.append(fmt_val(val))
            try:
                display_vals.append(f"{name}={val:.2f}°C")
            except TypeError:
                display_vals.append(f"{name}=NaN")

        # Write CSV (fatal if this fails)
        if self.csv_writer is not None:
            try:
                self.csv_writer.writerow(row)
                self.csv_file.flush()
            except Exception as e:
                messagebox.showerror("File error. Get Kailani.", f"Error writing to CSV:\n{e}")
                self.set_status("File error while writing CSV.", is_error=True)
                self.stop_logging(error=True)
                return

        self.last_line_var.set(ts + " | " + "  ".join(display_vals))

        # Update trends text
        self.update_channel_trends(temps)

        # Live graph
        if self.start_time is not None:
            elapsed = time.time() - self.start_time
        else:
            elapsed = 0.0

        if self.graph_window is not None and self.graph_window.winfo_exists():
            self.graph_window.add_sample(elapsed, temps)
        else:
            self.graph_window = None

        # Duration check
        if self.duration_seconds is not None and self.start_time is not None:
            if elapsed >= self.duration_seconds:
                self.stop_logging(error=False)
                return

        # Schedule next poll
        self.after(int(SAMPLE_INTERVAL * 1000), self.poll_once)

    # ---------- Trend detection ---------- #

    def update_channel_trends(self, temps: Dict[int, float]):
        """
        Decide if each channel is increasing / decreasing / stable
        (within self.trend_threshold °C) from last self.trend_window readings.

        Text format per channel:
            CH1: stable (avg 54.2 °C, Δ=+0.7 °C)
        """
        if not self.active_channels:
            return

        lines = [
            f"Channel temperature trends (last ~{self.trend_window} readings, "
            f"stable within ±{self.trend_threshold:.1f} °C):"
        ]

        for ch, name in self.active_channels:
            hlist = self.channel_history.setdefault(ch, [])

            val = temps.get(ch, None)
            # Try to parse numeric and append to history
            try:
                v = float(val)
                if math.isnan(v):
                    raise ValueError
                hlist.append(v)
                if len(hlist) > self.trend_window:
                    del hlist[:-self.trend_window]
            except Exception:
                # If no new valid value and no history, nothing to say yet
                if not hlist:
                    lines.append(f"  {name}: no data")
                    continue

            if len(hlist) < 2:
                lines.append(f"  {name}: no data")
                continue

            vmin = min(hlist)
            vmax = max(hlist)
            avg = sum(hlist) / len(hlist)
            delta = hlist[-1] - hlist[0]

            if (vmax - vmin) <= self.trend_threshold:
                trend = "stable"
            else:
                if delta > 0:
                    trend = "increasing"
                elif delta < 0:
                    trend = "decreasing"
                else:
                    trend = "stable"

            lines.append(f"  {name}: {trend} (avg {avg:.1f} °C, Δ={delta:+.1f} °C)")

        self.channel_trends_var.set("\n".join(lines))

    # ---------- Stop / close ---------- #

    def stop_logging(self, error: bool = False):
        if not self.is_logging:
            return

        self.is_logging = False
        self.start_button["state"] = "normal"
        self.stop_button["state"] = "disabled"

        try:
            if self.logger is not None:
                self.logger.close()
        except Exception:
            pass
        self.logger = None

        try:
            if self.csv_file is not None:
                self.csv_file.close()
        except Exception:
            pass
        self.csv_file = None
        self.csv_writer = None

        if not error and self.data_filename and HAVE_OPENPYXL:
            create_colored_excel(self.data_filename)

        if self.arduino is not None:
            try:
                self.arduino.close()
            except Exception:
                pass
            self.arduino = None

        self.set_status("Idle.")
        if not error and self.data_filename:
            messagebox.showinfo("Logging finished", f"Data saved to:\n{self.data_filename}")

    def on_stop(self):
        if self.is_logging:
            self.stop_logging(error=False)

    def on_close(self):
        if self.is_logging:
            if not messagebox.askyesno(
                "Quit",
                "Logging is still running. Stop and exit?"
            ):
                return
            self.stop_logging(error=True)

        if self.graph_window is not None and self.graph_window.winfo_exists():
            try:
                self.graph_window.destroy()
            except Exception:
                pass
            self.graph_window = None

        self.destroy()


if __name__ == "__main__":
    import traceback
    try:
        app = ThermalLoggerApp()
        app.mainloop()
    except Exception:
        traceback.print_exc()
        input("Error occurred, press Enter to exit...")


