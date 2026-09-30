@echo off
setlocal
cd /d "%~dp0"
set "PYTHON_EXE=D:\Anaconda\python.exe"
if not exist "%PYTHON_EXE%" (
  echo ERROR: Python was not found at %PYTHON_EXE%.
  pause
  exit /b 1
)
"%PYTHON_EXE%" -m streamlit run "%~dp0app.py"
if errorlevel 1 pause
