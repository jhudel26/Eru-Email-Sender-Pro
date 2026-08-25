@echo off
REM Eru Email Sender Pro - EXE Build Script
REM This script only builds the EXE (no installer)

echo ========================================
echo Eru Email Sender Pro - EXE Build Script
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher
    pause
    exit /b 1
)

echo Step 1: Installing dependencies...
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    echo Please check your Python version and pip installation
    echo Current Python version:
    python --version
    echo.
    echo If this is a version compatibility issue, check requirements.txt
    pause
    exit /b 1
)
echo Dependencies installed successfully.
echo.

echo Step 2: Building EXE with PyInstaller...
python -m PyInstaller --clean "Eru Email Sender Pro.spec"
if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed
    pause
    exit /b 1
)
echo.

echo Step 3: Checking if EXE was created...
if not exist "dist\Eru Email Sender Pro.exe" (
    echo ERROR: EXE file was not created
    pause
    exit /b 1
)
echo EXE created successfully: dist\Eru Email Sender Pro.exe
echo.

echo ========================================
echo EXE BUILD COMPLETED SUCCESSFULLY!
echo ========================================
echo.
echo Output file:
echo - EXE: dist\Eru Email Sender Pro.exe
echo.
echo You can now test the EXE by running:
echo cd dist
echo "Eru Email Sender Pro.exe"
echo.
pause