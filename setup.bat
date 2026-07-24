@echo off
title FNCL Group Manulife Investment Report Generator Installer
echo ==========================================================
echo       FNCL Group Manulife Investment Report Generator Setup Wizard
echo ==========================================================
echo.
echo Checking for Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not added to your system PATH.
    echo Please download and install Python 3.10 or 3.11 from https://www.python.org/
    echo Make sure to check the box "Add Python to PATH" during installation.
    echo.
    pause
    exit /b
)

echo [1/2] Creating Python Virtual Environment (venv)...
cd backend
if exist venv (
    echo Virtual environment already exists.
) else (
    python -m venv venv
)
if %errorlevel% neq 0 (
    echo [ERROR] Failed to create virtual environment.
    cd ..
    pause
    exit /b
)

echo [2/2] Installing dependencies from requirements.txt...
venv\Scripts\pip.exe install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    cd ..
    pause
    exit /b
)

echo Setup complete!
cd ..
echo.
echo ==========================================================
echo Installation Successful! 
echo You can now run the software by double-clicking 'run.bat'.
echo ==========================================================
echo.
pause
