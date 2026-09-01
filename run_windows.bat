@echo off
REM Daily use, once the environment is already set up (does not reinstall dependencies).
REM If this errors saying .venv is missing, run setup_windows.bat first.
"%~dp0.venv\Scripts\python.exe" "%~dp0main.py"
if errorlevel 1 pause
