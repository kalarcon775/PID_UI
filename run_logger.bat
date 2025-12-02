@echo off
cd /d "C:\Users\engsupport\OneDrive - Lux Dynamics\Documents\omega_pid_ui"

REM Use the 32-bit Python that works with the Omega DLL
"C:\Users\engsupport\AppData\Local\Programs\Python\Python313-32\python.exe" main_logger.py

echo.
echo Logging finished. Press any key to close this window...
pause >nul
