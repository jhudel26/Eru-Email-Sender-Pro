@echo off
REM Eru Email Sender Pro - Build Script
REM This script builds the EXE and then creates the installer

echo ========================================
echo Eru Email Sender Pro - Build Script
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
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)
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

echo Step 4: Creating installer with Inno Setup...
REM Check if Inno Setup is installed
set INNO_PATH="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %INNO_PATH% (
    set INNO_PATH="C:\Program Files\Inno Setup 6\ISCC.exe"
)
if not exist %INNO_PATH% (
    set INNO_PATH="C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
)
if not exist %INNO_PATH% (
    set INNO_PATH="C:\Program Files\Inno Setup 5\ISCC.exe"
)

if not exist %INNO_PATH% (
    echo ERROR: Inno Setup is not installed
    echo Please install Inno Setup from https://jrsoftware.org/isdl.php
    pause
    exit /b 1
)

%INNO_PATH% installer_script.iss
if %errorlevel% neq 0 (
    echo ERROR: Inno Setup build failed
    pause
    exit /b 1
)
echo.

echo ========================================
echo BUILD COMPLETED SUCCESSFULLY!
echo ========================================
echo.
echo Output files:
echo - EXE: dist\Eru Email Sender Pro.exe
echo - Installer: installer_output\Eru Email Sender Pro-Setup-2.0.0.exe
echo.
echo You can now distribute the installer file.
echo.
pause