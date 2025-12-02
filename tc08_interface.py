# tc08_interface.py
import ctypes
from ctypes import c_int16, c_int8, c_float, POINTER
from typing import Dict, List

DLL_PATH = r"C:\Program Files (x86)\Omega\Data Logging Software\usbtc08.dll"

class TC08Interface:
    """
    Simple wrapper around usbtc08.dll using usb_tc08_get_single.
    - Opens the device on creation
    - Returns one full set of temps (ch0..ch8) on read()
    """

    def __init__(self, dll_path: str = DLL_PATH, mains: int = 0, tc_type: str = "K"):
        self._dll = ctypes.CDLL(dll_path)

        # Declare the few functions we use
        self._dll.usb_tc08_open_unit.restype = c_int16

        self._dll.usb_tc08_set_mains.argtypes = [c_int16, c_int16]
        self._dll.usb_tc08_set_mains.restype = c_int16

        self._dll.usb_tc08_set_channel.argtypes = [c_int16, c_int16, c_int8]
        self._dll.usb_tc08_set_channel.restype = c_int16

        self._dll.usb_tc08_get_single.argtypes = [
            c_int16,              # handle
            POINTER(c_float),     # temp[9]
            POINTER(c_int16),     # overflow flags
            c_int16               # units
        ]
        self._dll.usb_tc08_get_single.restype = c_int16

        self._dll.usb_tc08_close_unit.argtypes = [c_int16]
        self._dll.usb_tc08_close_unit.restype = None

        # Open device
        self.handle = self._dll.usb_tc08_open_unit()
        if self.handle <= 0:
            raise RuntimeError(f"TC-08 open failed, handle={self.handle}")

        # Mains rejection: 0 = 50 Hz, 1 = 60 Hz
        self._dll.usb_tc08_set_mains(self.handle, mains)

        # Enable all channels as given thermocouple type (0..8, 0=cold junction)
        tc_char = ord(tc_type.upper())
        for ch in range(0, 9):
            r = self._dll.usb_tc08_set_channel(self.handle, ch, tc_char)
            if r != 1:
                raise RuntimeError(f"set_channel({ch}) failed with code {r}")

        # Buffers reused every read()
        self._temp_array = (c_float * 9)()
        self._overflow = c_int16(0)
        self._units = c_int16(0)  # 0 = °C

    def read(self) -> Dict[int, float]:
        """
        Read one full set of temperatures.
        Returns a dict: {channel_number: temp_in_C}
        ch 0 = cold junction, 1–8 = thermocouples.
        """
        result = self._dll.usb_tc08_get_single(
            self.handle,
            self._temp_array,
            ctypes.byref(self._overflow),
            self._units,
        )
        if result != 1:
            raise RuntimeError(f"usb_tc08_get_single failed, code={result}")

        temps: Dict[int, float] = {}
        for ch in range(9):
            temps[ch] = float(self._temp_array[ch])
        return temps

    def close(self):
        if getattr(self, "handle", None) is not None and self.handle > 0:
            self._dll.usb_tc08_close_unit(self.handle)
            self.handle = -1

    def __del__(self):
        # Best-effort cleanup
        try:
            self.close()
        except Exception:
            pass
