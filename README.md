# PID_UI  

**GUI Thermal Logger for LUX Dynamics**

## 📄 Overview  
PID_UI is a desktop application (Tkinter-based) that automates thermal testing using a Pico TC-08 data logger — with optional ambient-temperature control and logging via an Arduino. It logs channel data with timestamps, saves to CSV, and automatically generates a color-coded Excel (.xlsx) file for easy review and post-processing. It supports custom metadata, user-defined channel names, optional timed runs, and ambient setpoint control.

## 💡 Why this exists  
- Simplifies repetitive thermal test logging.  
- Combines multi-channel TC-08 measurements with Arduino ambient control in one interface.  
- Produces human-readable, formatted Excel output for quick analysis or sharing.  
- Useful for long-duration tests where manual logging would be error-prone.

## 🛠️ Features  
- User-friendly GUI for configuring tests (test name, tester, fixture, notes, channels).  
- Supports 0–8 TC-08 inputs + optional cold-junction sensor.  
- Optional Arduino ambient control: set a target temperature, log ambient temperature + PWM.  
- CSV logging with timestamp + channel readings (and ambient if enabled).  
- Automatic export to color-coded Excel for better readability.  
- Optional duration (run for N minutes) or unlimited logging.  
- Configurable channel names and metadata fields for traceability.  

## 🎯 System Requirements  
- Python 3.x  
- Packages: `tkinter`, `openpyxl` (for Excel export)  
- Hardware: Pico TC-08 logger (with `tc08_interface.py`), optional Arduino + ambient-control circuit (via serial/pyserial)  

## 🚀 Installation & Usage  

```bash
git clone https://github.com/kalarcon775/PID_UI.git
cd PID_UI
pip install openpyxl pyserial   # optional, only if using Excel export and Arduino
python main_logger.py            # or use run_logger.bat on Windows
