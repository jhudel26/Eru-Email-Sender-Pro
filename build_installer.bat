@echo off
echo Building Eru Email Sender Pro Installer...
echo.

echo Step 1: Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist installer_output rmdir /s /q installer_output

echo Step 2: Building executable with PyInstaller...
python -m PyInstaller --clean "Eru Email Sender Pro.spec"
if %ERRORLEVEL% neq 0 (
    echo ERROR: PyInstaller build failed!
    pause
    exit /b 1
)

echo Step 3: Creating installer with Inno Setup...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer_script.iss
if %ERRORLEVEL% neq 0 (
    echo ERROR: Inno Setup build failed!
    pause
    exit /b 1
)

echo.
echo Build completed successfully!
echo Installer location: installer_output\EruEmailSenderPro_Setup.exe
echo.
pause
