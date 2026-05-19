# logger_core.py
"""
Core logging interfaces for the TC-08 thermometer and Arduino ambient controller.

This module defines shared configuration constants used by the UI and graph.
The Arduino no longer needs its own temperature sensor. Python sends the TC-08
ambient thermocouple reading to the Arduino using AMB:<temp>.
"""

import time
import math
from typing import Optional, Tuple

# ---------------- Configuration Constants ---------------- #

TREND_WINDOW_DEFAULT = 10
TREND_THRESHOLD_DEFAULT = 3.0
SAMPLE_INTERVAL = 1.0
MAX_GRAPH_POINTS = 2000

# ---------------- TC-08 Interface ---------------- #

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
        Open the Arduino serial port.

        The Arduino receives:
            SET:<setpoint_C>
            AMB:<ambient_temp_C>

        The Arduino may return:
            AMB:<ambient_C>,HOLD:<setpoint_C>,PWM:<pwm>,STATUS:<status>
        """

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
        """
        Send one newline-terminated command to the Arduino.
        """
        try:
            line = text.strip() + "\n"
            self.ser.write(line.encode("ascii"))
        except Exception:
            pass

    def set_hold(self, temp_c: float) -> None:
        """
        Send a setpoint to the Arduino.

        Example:
            SET:40.00
        """
        try:
            temp_c = float(temp_c)
            if math.isnan(temp_c):
                return

            self.latest_hold = temp_c
            self._write_line(f"SET:{temp_c:.2f}")

        except Exception:
            pass

    def send_ambient(self, temp_c: float) -> None:
        """
        Send the TC-08 ambient thermocouple reading to the Arduino.

        Example:
            AMB:38.42
        """
        try:
            temp_c = float(temp_c)
            if math.isnan(temp_c):
                return

            self.latest_ambient = temp_c
            self._write_line(f"AMB:{temp_c:.2f}")

        except Exception:
            pass

    def poll(self) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[str]]:
        """
        Read any pending serial lines from the Arduino.

        Returns:
            latest_ambient, latest_hold, latest_pwm, latest_status

        The expected Arduino format is:
            AMB:38.42,HOLD:40.00,PWM:155,STATUS:OK

        It also supports old TEMP format:
            TEMP:38.42,HOLD:40.00,PWM:155
        """

        line = None

        try:
            while self.ser.in_waiting:
                raw = self.ser.readline()
                if not raw:
                    break

                decoded = raw.decode("ascii", errors="ignore").strip()

                if decoded:
                    line = decoded

        except Exception:
            return (
                self.latest_ambient,
                self.latest_hold,
                self.latest_pwm,
                self.latest_status,
            )

        if not line:
            return (
                self.latest_ambient,
                self.latest_hold,
                self.latest_pwm,
                self.latest_status,
            )

        try:
            parts = [p.strip() for p in line.split(",")]

            for p in parts:
                if p.startswith("AMB:"):
                    self.latest_ambient = float(p.split("AMB:", 1)[1])

                elif p.startswith("TEMP:"):
                    self.latest_ambient = float(p.split("TEMP:", 1)[1])

                elif p.startswith("HOLD:"):
                    self.latest_hold = float(p.split("HOLD:", 1)[1])

                elif p.startswith("PWM:"):
                    self.latest_pwm = float(p.split("PWM:", 1)[1])

                elif p.startswith("STATUS:"):
                    self.latest_status = p.split("STATUS:", 1)[1]

                elif p.startswith("ERR:"):
                    self.latest_status = p

        except ValueError:
            pass

        return (
            self.latest_ambient,
            self.latest_hold,
            self.latest_pwm,
            self.latest_status,
        )

    def close(self) -> None:
        """
        Turn the setpoint down and close the serial port.
        """
        try:
            self._write_line("SET:0")
            time.sleep(0.05)
        except Exception:
            pass

        try:
            self.ser.close()
        except Exception:
            pass
