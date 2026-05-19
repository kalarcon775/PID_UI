# logger_core.py
"""
Core logging interfaces for the TC-08 thermometer and Arduino ambient controller.

This module also defines shared configuration constants used by the UI and graph.
"""

import time
import math
from typing import Optional

# ---------------- Configuration Constants ---------------- #

TREND_WINDOW_DEFAULT = 10         # Default # of recent samples to compute channel trends
TREND_THRESHOLD_DEFAULT = 3.0     # Default temperature band [°C] considered "stable"
SAMPLE_INTERVAL = 1.0             # Default logging interval [seconds] between TC-08 reads
MAX_GRAPH_POINTS = 2000           # Max samples per channel stored for the live graph

# ---------------- TC-08 Interface ---------------- #
# The actual implementation is provided in tc08_interface.py
# and imported here so the UI can just do: `from logger_core import TC08Interface`.

from tc08_interface import TC08Interface  # type: ignore


# ---------------- Arduino Interface ---------------- #

try:
    import serial  # type: ignore
    HAVE_SERIAL = True
except ImportError:
    serial = None
    HAVE_SERIAL = False


class ArduinoInterface:
    def __init__(self, port: str, baudrate: int = 9600):
        global HAVE_SERIAL, serial

        if not HAVE_SERIAL:
            try:
                import serial as serial_mod  # type: ignore
                serial = serial_mod
                HAVE_SERIAL = True
            except ImportError:
                raise RuntimeError(
                    "pyserial not installed; cannot use ArduinoInterface. Get Kailani."
                )

        self.ser = serial.Serial(port, baudrate=baudrate, timeout=0.1)
        time.sleep(2.0)
        self.ser.reset_input_buffer()

        self.latest_ambient: Optional[float] = None
        self.latest_hold: Optional[float] = None
        self.latest_pwm: Optional[float] = None
        self.latest_status: Optional[str] = None

    def _write_line(self, text: str) -> None:
        try:
            self.ser.write((text.strip() + "\n").encode("ascii"))
        except Exception:
            pass

    def set_hold(self, temp_c: float) -> None:
        self.latest_hold = temp_c
        self._write_line(f"SET:{temp_c:.2f}")

    def send_ambient(self, temp_c: float) -> None:
        try:
            if temp_c is None or math.isnan(float(temp_c)):
                return
            self.latest_ambient = float(temp_c)
            self._write_line(f"AMB:{float(temp_c):.2f}")
        except Exception:
            pass

    def poll(self):
        line = None

        try:
            while self.ser.in_waiting:
                raw = self.ser.readline()
                if not raw:
                    break
                line = raw.decode("ascii", errors="ignore").strip()
        except Exception:
            return self.latest_ambient, self.latest_hold, self.latest_pwm, self.latest_status

        if not line:
            return self.latest_ambient, self.latest_hold, self.latest_pwm, self.latest_status

        try:
            parts = [p.strip() for p in line.split(",")]
            for p in parts:
                if p.startswith("AMB:"):
                    self.latest_ambient = float(p.split("AMB:")[1])
                elif p.startswith("TEMP:"):
                    self.latest_ambient = float(p.split("TEMP:")[1])
                elif p.startswith("HOLD:"):
                    self.latest_hold = float(p.split("HOLD:")[1])
                elif p.startswith("PWM:"):
                    self.latest_pwm = float(p.split("PWM:")[1])
                elif p.startswith("STATUS:"):
                    self.latest_status = p.split("STATUS:")[1]
                elif p.startswith("ERR:"):
                    self.latest_status = p
        except ValueError:
            pass

        return self.latest_ambient, self.latest_hold, self.latest_pwm, self.latest_status

    def close(self) -> None:
        try:
            self._write_line("SET:0")
            self.ser.close()
        except Exception:
            pass
