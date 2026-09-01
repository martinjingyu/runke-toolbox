@echo off
REM Double-click this file: first run auto-installs Python + dependencies, then starts the app.
REM After that, use run_windows.bat to open the app without reinstalling anything.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_windows.ps1"
pause
