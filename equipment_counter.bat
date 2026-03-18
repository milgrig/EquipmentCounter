@echo off
chcp 65001 >nul 2>&1
title Equipment Counter

:: Change to script directory
cd /d "%~dp0"

:: Find Python (try common locations)
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
if not defined PYTHON (
    where python3 >nul 2>&1 && set PYTHON=python3
)
if not defined PYTHON (
    for /f "delims=" %%i in ('dir /b /s "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python3*\python.exe" 2^>nul') do (
        set "PYTHON=%%i"
    )
)
if not defined PYTHON (
    echo.
    echo  Python not found. Please install Python 3.10+
    echo  https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo Using: %PYTHON%

:: Check and install dependencies
%PYTHON% -c "import ezdxf; import pdfplumber" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing dependencies...
    %PYTHON% -m pip install -r requirements.txt
    echo.
)

:: Launch GUI
%PYTHON% equipment_gui.py
if %errorlevel% neq 0 (
    echo.
    echo  Error launching GUI.
    pause
)
