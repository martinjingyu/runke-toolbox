@echo off
REM Resets this machine back to "nothing installed" so you can retest setup_windows.bat
REM from scratch. Removes .venv, uninstalls Python/Tesseract if this project installed
REM them, and cleans up the PATH entries that were added.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_windows.ps1"
pause
