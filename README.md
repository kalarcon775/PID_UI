# Thermal Logger

GUI-based thermal test logger for LUX Dynamics.

## Overview

Thermal Logger is a desktop application built with Python and Tkinter for automating thermal testing with a Pico TC-08 data logger. It records multi-channel temperature data, saves results to CSV, and can automatically generate a formatted Excel file for easier review and post-processing.

The program also supports optional ambient-temperature control and logging through an Arduino. Users can configure test metadata, rename channels, choose logging duration, and include fixture or test notes for traceability.

## Why This Exists

Thermal testing often requires repeated manual measurements, long-duration logging, and organized data review. Thermal Logger simplifies that workflow by combining data collection, metadata entry, optional ambient control, and formatted output generation in one interface.

This tool was created to make thermal validation testing more consistent, readable, and easier to share across engineering and product development work.

## Features

- Tkinter-based graphical user interface
- Supports up to 8 Pico TC-08 temperature channels
- Optional cold-junction sensor logging
- Optional Arduino-based ambient temperature control
- Ambient temperature and PWM logging when enabled
- User-defined test name, tester, fixture, notes, and channel names
- CSV output with sample number and channel readings
- Automatic Excel export with formatted, color-coded columns
- Optional timed test runs
- Unlimited logging mode for long-duration tests
- Metadata fields for better test traceability

## System Requirements

### Software

- Python 3.x
- Tkinter
- openpyxl, used for Excel export
- pyserial, required only if using Arduino-based ambient control

### Hardware

- Pico TC-08 data logger
- Thermocouples connected to TC-08 channels
- Optional Arduino and ambient-control circuit

## Installation

Clone the repository:

```bash
git clone https://github.com/kalarcon775/PID_UI.git
cd PID_UI
```

Install the required Python packages:

```bash
pip install openpyxl pyserial
```

`openpyxl` is used to generate formatted Excel files. `pyserial` is only required when using Arduino-based ambient-temperature control.

## Usage

Run the application:

```bash
python main_logger.py
```

On Windows, you may also run:

```bash
run_logger.bat
```

In the GUI, enter the test information, select the active channels, name the channels if needed, add fixture details or notes, and choose whether the test should run for a set duration or continue until stopped manually.

After configuration, click **Start Logging**. The program will save the logged data to the selected output folder.

## Output Files

Thermal Logger generates a CSV file containing:

- Sample number
- Logged TC-08 temperature channels
- Optional cold-junction reading
- Optional ambient temperature
- Optional PWM value

If Excel export is available, the program also creates an `.xlsx` file with the same logged data. The Excel file includes formatted columns, color coding, and grid borders to make the data easier to review, analyze, and share.

## Project Structure

```text
PID_UI/
├── main_logger.py
├── tc08_interface.py
├── run_logger.bat
├── README.md
└── output files
```

## Notes

This project is intended for thermal validation and product development testing. It is especially useful for long-duration thermal tests where manual logging can be repetitive, inconsistent, or error-prone.

The Arduino ambient-control feature is optional. The logger can still be used for TC-08 temperature data collection without the Arduino system enabled.

## Credits

Created by Kailani Puava Alarcon at LUX Dynamics.
