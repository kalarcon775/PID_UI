# graph_window.py
"""
Live graph window for the LUX Thermal Logger.

Shows time vs temperature for multiple channels, with:
  - Zoom in/out (time window in minutes)
  - Slider to scroll earlier/later in time
  - Small y-axis with ticks + numeric labels
  - Grid lines for readability
  - Channel visibility checkboxes
  - Hover readout showing nearest time + visible channel values
"""

import math
from typing import Dict, List, Tuple

import tkinter as tk
from tkinter import ttk

from logger_core import MAX_GRAPH_POINTS


class LiveGraphWindow(tk.Toplevel):
    """
    Separate window that shows a live graph of time vs temperature
    for all active TC-08 channels.
    """

    def __init__(self, master):
        super().__init__(master)
        self.title("Live Temperature Graph")
        self.geometry("950x500")

        # history[ch] = {"t": [elapsed_s...], "v": [temp_C...]}
        self.history: Dict[int, Dict[str, List[float]]] = {}
        self.active_channels: List[Tuple[int, str]] = []

        self.window_sec = 300.0  # default 5 minute visible window
        self.max_points = MAX_GRAPH_POINTS

        self.graph_colors = [
            "blue", "red", "green", "purple",
            "orange", "brown", "magenta", "cyan"
        ]

        self.pan_var = tk.DoubleVar(value=0.0)

        # axis / plot geometry state for hover
        self.plot_left = None
        self.plot_right = None
        self.plot_top = None
        self.plot_bottom = None
        self.tmin = None
        self.tmax = None
        self.vmin = None
        self.vmax = None

        # channel visibility flags
        self.channel_visibility: Dict[int, tk.BooleanVar] = {}

        self._build_ui()
        self._update_window_label()

    # ---------------- UI setup ---------------- #

    def _build_ui(self):
        # Top controls: zoom + slider
        controls = ttk.Frame(self, padding=8)
        controls.pack(fill="x")

        self.window_label_var = tk.StringVar()
        ttk.Button(controls, text="Zoom -", command=self.zoom_out).pack(side="left")
        ttk.Button(controls, text="Zoom +", command=self.zoom_in).pack(side="left", padx=(2, 8))
        ttk.Label(controls, textvariable=self.window_label_var).pack(side="left")

        # Pretty slider
        style = ttk.Style(self)
        style.configure(
            "Pan.Horizontal.TScale",
            troughcolor="#e5e5e5",
        )

        slider_frame = ttk.Frame(controls)
        slider_frame.pack(side="right")
        ttk.Label(slider_frame, text="Earlier").pack(side="left", padx=(0, 4))
        self.pan_scale = ttk.Scale(
            slider_frame,
            from_=0.0,
            to=100.0,
            orient="horizontal",
            variable=self.pan_var,
            command=lambda v: self.redraw(),
            style="Pan.Horizontal.TScale",
            length=260,
        )
        self.pan_scale.pack(side="left")
        ttk.Label(slider_frame, text="Later").pack(side="left", padx=(4, 0))

        # Channel toggles row
        self.toggle_frame = ttk.Frame(self, padding=(8, 0))
        self.toggle_frame.pack(fill="x", anchor="w")

        # Main drawing canvas
        self.canvas = tk.Canvas(self, bg="white")
        self.canvas.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.canvas.bind("<Leave>", lambda e: self.hover_label_var.set(""))

        # Hover readout
        self.hover_label_var = tk.StringVar(value="Hover over plot for values")
        self.hover_label = ttk.Label(self, textvariable=self.hover_label_var, padding=(8, 2))
        self.hover_label.pack(fill="x", side="bottom", anchor="w")

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        """Just close this window; logger keeps running."""
        self.destroy()

    def _update_window_label(self):
        if self.window_sec is None:
            text = "Window: full"
        else:
            text = f"Window: {self.window_sec / 60.0:.2f} min"
        self.window_label_var.set(text)

    # ---------------- Channel configuration ---------------- #

    def set_channels(self, active_channels: List[Tuple[int, str]]):
        """
        Called by the main UI when a new run starts.
        """
        self.active_channels = list(active_channels)
        self.history.clear()
        self.refresh_channel_toggles()

    def refresh_channel_toggles(self):
        # Clear old checkboxes
        for child in self.toggle_frame.winfo_children():
            child.destroy()

        # Build new checkboxes for each channel
        for ch, name in self.active_channels:
            var = self.channel_visibility.get(ch)
            if var is None:
                var = tk.BooleanVar(value=True)
                self.channel_visibility[ch] = var
            cb = ttk.Checkbutton(
                self.toggle_frame,
                text=name,
                variable=var,
                command=self.redraw
            )
            cb.pack(side="left", padx=(0, 8))

    # ---------------- Data update ---------------- #

    def add_sample(self, elapsed: float, temps: Dict[int, float]):
        """
        Append a new reading for each channel and redraw the graph.
        `elapsed` is seconds since the run started.
        """
        for ch, _name in self.active_channels:
            val = temps.get(ch, float("nan"))
            try:
                if val is None or math.isnan(float(val)):
                    continue
            except (TypeError, ValueError):
                continue

            if ch not in self.history:
                self.history[ch] = {"t": [], "v": []}

            self.history[ch]["t"].append(float(elapsed))
            self.history[ch]["v"].append(float(val))

            # Limit memory
            if len(self.history[ch]["t"]) > self.max_points:
                self.history[ch]["t"] = self.history[ch]["t"][-self.max_points:]
                self.history[ch]["v"] = self.history[ch]["v"][-self.max_points:]

        self.redraw()

    # ---------------- Zoom / pan ---------------- #

    def zoom_in(self):
        """Halve the visible time window, down to a minimum."""
        if self.window_sec is None:
            self.window_sec = 300.0
        self.window_sec = max(5.0, self.window_sec / 2.0)
        self._update_window_label()
        self.redraw()

    def zoom_out(self):
        """Double the visible time window, up to the full history."""
        if not self.history:
            return

        all_times: List[float] = []
        for h in self.history.values():
            all_times.extend(h["t"])
        if not all_times:
            return

        total_span = max(all_times) - min(all_times)
        if total_span <= 0:
            return

        if self.window_sec is None:
            self.window_sec = total_span

        self.window_sec *= 2.0
        if self.window_sec >= total_span:
            self.window_sec = None  # show full range

        self._update_window_label()
        self.redraw()

    # ---------------- Drawing ---------------- #

    def redraw(self):
        """
        Redraw the entire graph based on current history and zoom/pan/window settings.
        """
        # Reset hover geometry
        self.plot_left = self.plot_right = self.plot_top = self.plot_bottom = None
        self.tmin = self.tmax = self.vmin = self.vmax = None

        if not self.history:
            self.canvas.delete("all")
            return

        # Collect global time/value range
        all_times: List[float] = []
        all_vals: List[float] = []
        for h in self.history.values():
            all_times.extend(h["t"])
            all_vals.extend(h["v"])

        if len(all_times) < 2 or not all_vals:
            self.canvas.delete("all")
            return

        global_tmin = min(all_times)
        global_tmax = max(all_times)

        # Determine displayed tmin/tmax
        if self.window_sec is None:
            tmin = global_tmin
            tmax = global_tmax
        else:
            window = self.window_sec
            span = max(global_tmax - global_tmin, 1e-6)
            window = min(window, span)
            start_min = global_tmin
            start_max = global_tmax - window
            if start_max <= start_min:
                tmin = global_tmin
            else:
                frac = max(0.0, min(1.0, self.pan_var.get() / 100.0))
                tmin = start_min + frac * (start_max - start_min)
            tmax = tmin + window

        if tmax <= tmin:
            tmax = tmin + 1.0

        # y-limits auto
        vmin = min(all_vals)
        vmax = max(all_vals)
        if vmax <= vmin:
            vmax = vmin + 1.0

        w = self.canvas.winfo_width() or 400
        h = self.canvas.winfo_height() or 300

        legend_width = 130
        margin_right = 40
        margin_top = 20
        margin_bottom = 30
        margin_left_extra = 30  # space for y labels

        plot_left = legend_width + margin_left_extra
        plot_right = w - margin_right
        plot_top = margin_top
        plot_bottom = h - margin_bottom

        if plot_right <= plot_left + 10:
            self.canvas.delete("all")
            return

        # Store geometry for hover calculations
        self.plot_left = plot_left
        self.plot_right = plot_right
        self.plot_top = plot_top
        self.plot_bottom = plot_bottom
        self.tmin = tmin
        self.tmax = tmax
        self.vmin = vmin
        self.vmax = vmax

        self.canvas.delete("all")

        # Grid + axis labels
        n_x = 5
        n_y = 5
        time_span = tmax - tmin

        # Vertical grid lines + x labels (time in minutes)
        for i in range(n_x + 1):
            gx = plot_left + i * (plot_right - plot_left) / n_x
            self.canvas.create_line(
                gx, plot_top, gx, plot_bottom,
                fill="#e0e0e0"
            )
            t_here = tmin + time_span * i / n_x
            label = f"{t_here / 60.0:.1f}"  # minutes
            self.canvas.create_text(
                gx, plot_bottom + 12,
                text=label,
                anchor="n"
            )

        # Horizontal grid lines + y labels (temperature)
        for j in range(n_y + 1):
            gy = plot_top + j * (plot_bottom - plot_top) / n_y
            self.canvas.create_line(
                plot_left, gy, plot_right, gy,
                fill="#e0e0e0"
            )
            v_here = vmax - (vmax - vmin) * j / n_y
            self.canvas.create_line(
                plot_left - 4, gy, plot_left, gy,
                fill="black"
            )
            self.canvas.create_text(
                plot_left - 6, gy,
                text=f"{v_here:.1f}",
                anchor="e"
            )

        # Axes
        self.canvas.create_line(
            plot_left, plot_bottom, plot_right, plot_bottom,
            fill="black", width=2
        )
        self.canvas.create_line(
            plot_left, plot_top, plot_left, plot_bottom,
            fill="black", width=2
        )

        # Plot each visible channel
        for idx, (ch, name) in enumerate(self.active_channels):
            var = self.channel_visibility.get(ch)
            if var is not None and not var.get():
                continue

            if ch not in self.history:
                continue
            t_list = self.history[ch]["t"]
            v_list = self.history[ch]["v"]
            if len(t_list) < 2:
                continue

            coords: List[float] = []
            for t_val, v_val in zip(t_list, v_list):
                if t_val < tmin or t_val > tmax:
                    continue
                x = plot_left + (t_val - tmin) / (tmax - tmin) * (plot_right - plot_left)
                y = plot_bottom - (v_val - vmin) / (vmax - vmin) * (plot_bottom - plot_top)
                coords.extend((x, y))

            if len(coords) < 4:
                continue

            color = self.graph_colors[idx % len(self.graph_colors)]
            self.canvas.create_line(*coords, fill=color, width=2)

        # X-axis label
        window_min = (tmax - tmin) / 60.0
        self.canvas.create_text(
            (plot_left + plot_right) / 2,
            h - 10,
            text=f"Time (min) – window ≈ {window_min:.2f} min",
            anchor="center"
        )

        # Legend (only visible channels)
        legend_x = 10
        legend_y = 25
        for idx, (ch, name) in enumerate(self.active_channels):
            var = self.channel_visibility.get(ch)
            if var is not None and not var.get():
                continue
            color = self.graph_colors[idx % len(self.graph_colors)]
            self.canvas.create_rectangle(
                legend_x,
                legend_y - 5,
                legend_x + 20,
                legend_y + 5,
                fill=color,
                outline=color
            )
            self.canvas.create_text(
                legend_x + 25,
                legend_y,
                text=name,
                anchor="w"
            )
            legend_y += 18

    # ---------------- Hover readout ---------------- #

    def on_mouse_move(self, event):
        """
        Show nearest time (minutes) and value of each visible channel
        at the cursor's x-position.
        """
        if not self.history:
            self.hover_label_var.set("")
            return

        if (
            self.plot_left is None or self.plot_right is None or
            self.plot_top is None or self.plot_bottom is None or
            self.tmin is None or self.tmax is None
        ):
            self.hover_label_var.set("")
            return

        x = event.x
        y = event.y
        if x < self.plot_left or x > self.plot_right or y < self.plot_top or y > self.plot_bottom:
            self.hover_label_var.set("")
            return

        if self.plot_right == self.plot_left or self.tmax == self.tmin:
            self.hover_label_var.set("")
            return

        t_current = self.tmin + (x - self.plot_left) / (self.plot_right - self.plot_left) * (self.tmax - self.tmin)

        info = []
        for ch, name in self.active_channels:
            var = self.channel_visibility.get(ch)
            if var is not None and not var.get():
                continue
            if ch not in self.history:
                continue
            t_list = self.history[ch]["t"]
            v_list = self.history[ch]["v"]
            if not t_list:
                continue

            best_idx = None
            best_dt = None
            for idx, t_val in enumerate(t_list):
                if t_val < self.tmin or t_val > self.tmax:
                    continue
                dt = abs(t_val - t_current)
                if best_idx is None or dt < best_dt:
                    best_idx = idx
                    best_dt = dt
            if best_idx is None:
                continue
            t_val = t_list[best_idx]
            v_val = v_list[best_idx]
            info.append((name, t_val, v_val))

        if not info:
            self.hover_label_var.set("")
            return

        t_display = info[0][1] / 60.0  # minutes
        parts = [f"t={t_display:.2f} min"]
        for name, _, v in info:
            parts.append(f"{name}={v:.2f}°C")
        self.hover_label_var.set(" | ".join(parts))
