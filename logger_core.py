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
        """
        Open the given serial COM port at the specified baudrate.
        Raises RuntimeError if pyserial is not installed.
        """
        global HAVE_SERIAL, serial

        # Lazy re-check in case pyserial was installed after import
        if not HAVE_SERIAL:
            try:
                import serial as serial_mod  # type: ignore
                serial = serial_mod
                HAVE_SERIAL = True
            except ImportError:
                raise RuntimeError(
                    "pyserial not installed; cannot use ArduinoInterface. Get Kailani."
                )

        # timeout kept short so reads don't stall the GUI
        self.ser = serial.Serial(port, baudrate=baudrate, timeout=0.1)
        # Give the Arduino time to reset
        time.sleep(2.0)
        self.ser.reset_input_buffer()

        self.latest_temp: Optional[float] = None
        self.latest_hold: Optional[float] = None
        self.latest_pwm: Optional[float] = None

    def set_hold(self, temp_c: float) -> None:
        """
        Send a new ambient setpoint to the Arduino, e.g. SET:25.00\n
        """
        cmd = f"SET:{temp_c:.2f}\n"
        try:
            self.ser.write(cmd.encode("ascii"))
        except Exception:
            # Non-fatal: ignore write problems; caller can decide how to handle.
            pass

    def poll(self):
        """
        Read any pending lines from the serial buffer.

        Returns the latest parsed (temp_C, hold_C, pwm) tuple.
        If no new valid line is seen, returns the last known values.
        """
        line = None

        try:
            # Drain buffer to get the most recent complete line
            while self.ser.in_waiting:
                raw = self.ser.readline()
                if not raw:
                    break
                line = raw.decode("ascii", errors="ignore").strip()
        except Exception:
            # On read or decode problems, just return the last good values
            return self.latest_temp, self.latest_hold, self.latest_pwm

        if not line:
            return self.latest_temp, self.latest_hold, self.latest_pwm

        # Expected format: TEMP:25.30,HOLD:53.60,PWM:255
        try:
            if "TEMP:" in line:
                parts = [p.strip() for p in line.split(",")]
                for p in parts:
                    if p.startswith("TEMP:"):
                        self.latest_temp = float(p.split("TEMP:")[1])
                    elif p.startswith("HOLD:"):
                        self.latest_hold = float(p.split("HOLD:")[1])
                    elif p.startswith("PWM:"):
                        self.latest_pwm = float(p.split("PWM:")[1])
            else:
                # Fallback: just a bare temperature number
                self.latest_temp = float(line)
        except ValueError:
            # Ignore malformed lines, keep last good values
            pass

        return self.latest_temp, self.latest_hold, self.latest_pwm

    def close(self) -> None:
        """Close the serial port, ignoring any errors."""
        try:
            self.ser.close()
        except Exception:
            pass
