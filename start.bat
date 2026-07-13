@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ==================================================
echo   DPS Aircon Demand Simulation App
echo ==================================================
echo.

set "PYCMD="
where py >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD (
  where python >nul 2>nul && set "PYCMD=python"
)
if not defined PYCMD (
  echo [ERROR] Python not found.
  echo Please install Python 3.9 or newer first:
  echo   https://www.python.org/downloads/
  echo During install, check "Add Python to PATH".
  echo.
  pause
  exit /b 1
)

%PYCMD% -c "import streamlit" >nul 2>nul
if errorlevel 1 (
  echo First-time setup: installing required libraries. This may take a few minutes...
  echo.
  %PYCMD% -m pip install --upgrade pip
  %PYCMD% -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo [ERROR] Library install failed. Please check your internet connection.
    pause
    exit /b 1
  )
  echo.
  echo Setup complete.
  echo.
)

echo Starting the app. Your browser will open automatically.
echo   If it does not open, go to  http://localhost:8501
echo   To stop: press Ctrl+C in this window, or close it.
echo.
%PYCMD% -m streamlit run app.py

echo.
echo App stopped.
pause
