@echo off
echo Building Eru Email Sender Pro Installer...
echo.

REM Create icon from PNG (if you have one)
REM python -c "from PIL import Image; img = Image.open('app_icon.png'); img.save('icon.ico', format='ICO', sizes=[16, 32, 48, 64, 128, 256])"

REM Build the executable using PyInstaller spec file
pyinstaller --clean build.spec

echo.
echo Build completed! 
echo Executable location: dist\Eru Email Sender Pro.exe
echo.

REM Create installer directory structure
if not exist "installer" mkdir installer
if not exist "installer\files" mkdir installer\files

REM Copy executable to installer directory
copy "dist\Eru Email Sender Pro.exe" "installer\files\"

REM Copy additional files if needed
copy "README.md" "installer\files\" 2>nul
copy "LICENSE" "installer\files\" 2>nul

echo.
echo Files copied to installer directory.
echo You can now create an installer using NSIS or Inno Setup.
echo.

pause
