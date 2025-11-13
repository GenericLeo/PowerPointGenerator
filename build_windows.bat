@echo off
REM PowerPoint Generator - Windows Build Script
REM This script builds a standalone Windows executable

echo ======================================
echo PowerPoint Generator - Build Script
echo ======================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or later from python.org
    pause
    exit /b 1
)

REM Check if virtual environment exists
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
)

REM Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat

REM Install requirements
echo Installing requirements...
pip install -r requirements.txt

REM Install PyInstaller if not already installed
echo Checking for PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
) else (
    echo PyInstaller already installed
)

echo.
echo Building application...
echo.

REM Clean previous builds
if exist "dist" (
    echo Cleaning previous builds...
    rmdir /s /q dist
)

if exist "build" (
    rmdir /s /q build
)

REM Build the application
pyinstaller build_windows.spec

REM Check if build was successful
if exist "dist\PowerPointGenerator.exe" (
    echo.
    echo ======================================
    echo Build successful!
    echo ======================================
    echo.
    echo Your application is located at:
    echo   dist\PowerPointGenerator.exe
    echo.
    echo You can:
    echo   1. Double-click to run it
    echo   2. Create a shortcut on your desktop
    echo   3. Share the entire 'dist' folder with others
    echo.
    pause
) else (
    echo.
    echo ======================================
    echo Build failed
    echo ======================================
    echo.
    echo Please check the error messages above.
    pause
    exit /b 1
)
