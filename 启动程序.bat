@echo off
setlocal
cd /d "%~dp0"
set "PYTHONW_EXE=D:\Anaconda\pythonw.exe"
if not exist "%PYTHONW_EXE%" (
  echo ERROR: Python GUI runtime was not found at %PYTHONW_EXE%.
  pause
  exit /b 1
)
start "" "%PYTHONW_EXE%" "%~dp0desktop_app.py"
if errorlevel 1 (
  echo ERROR: The desktop app could not start.
  pause
)
