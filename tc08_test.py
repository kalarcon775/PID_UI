import ctypes
from ctypes import c_int16, c_int8, c_float, POINTER
import time

# Path to your Omega / TC-08 DLL
dll_path = r"C:\Program Files (x86)\Omega\Data Logging Software\usbtc08.dll"
usbtc08 = ctypes.CDLL(dll_path)

# --- Function prototypes ---

# open_unit
usbtc08.usb_tc08_open_unit.restype = c_int16

# set mains (50/60 Hz)
usbtc08.usb_tc08_set_mains.argtypes = [c_int16, c_int16]
usbtc08.usb_tc08_set_mains.restype  = c_int16

# set channel (thermocouple type)
usbtc08.usb_tc08_set_channel.argtypes = [c_int16, c_int16, c_int8]
usbtc08.usb_tc08_set_channel.restype  = c_int16

# get_single: one complete set of readings (all 9 channels) on demand
usbtc08.usb_tc08_get_single.argtypes = [
    c_int16,                # handle
    POINTER(c_float),       # temp[9]
    POINTER(c_int16),       # overflow_flags
    c_int16                 # units
]
usbtc08.usb_tc08_get_single.restype = c_int16

# close_unit
usbtc08.usb_tc08_close_unit.argtypes = [c_int16]
usbtc08.usb_tc08_close_unit.restype  = None


def main():
    # --- Open device ---
    handle = usbtc08.usb_tc08_open_unit()
    print("Handle:", handle)
    if handle <= 0:
        raise RuntimeError("Failed to open TC-08 (check USB and drivers).")

    # 0 = 50 Hz, 1 = 60 Hz mains rejection
    usbtc08.usb_tc08_set_mains(handle, 0)

    # Enable channels 0..8 as Type K
    for ch in range(0, 9):   # 0 = cold junction, 1–8 = thermocouples
        r = usbtc08.usb_tc08_set_channel(handle, ch, ord('K'))
        print(f"set_channel({ch}) -> {r}")

    # Allocate buffers for get_single
    temp_array      = (c_float * 9)()   # temps for channels 0..8
    overflow_flags  = c_int16(0)
    units           = c_int16(0)        # 0 = °C

    print("\nReading temperatures once per second (Ctrl+C to stop)...\n")

    try:
        while True:
            result = usbtc08.usb_tc08_get_single(
                handle,
                temp_array,
                ctypes.byref(overflow_flags),
                units
            )

            if result != 1:
                print("usb_tc08_get_single failed, result =", result)
            else:
                # Convert the C array to a plain Python list
                temps = [float(temp_array[i]) for i in range(9)]
                # ch0 = cold junction, ch1–8 = thermocouples
                print("Temps (°C):", temps)

            time.sleep(1.0)

    except KeyboardInterrupt:
        print("\nStopping...")

    finally:
        usbtc08.usb_tc08_close_unit(handle)
        print("Device closed.")


if __name__ == "__main__":
    main()
