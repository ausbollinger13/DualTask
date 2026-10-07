@echo off
REM Build DualTask as a single Windows executable
REM Requires: pip install pyinstaller

pyinstaller --onefile --windowed --name DualTask dual_task.py
echo.
echo Build complete. Executable is in the dist\ folder.
pause
