@echo off
REM Build script for creating .exe file from Python application
REM Usage: build_exe.bat
REM Version is set in converter.spec file

echo Building Excel Time Tracker Converter executable...

REM Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Install/upgrade dependencies
echo Installing dependencies...
python -m pip install --upgrade pip
pip install -r requirements.txt

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Failed to install dependencies!
    exit /b 1
)

REM Build executable using PyInstaller
echo.
echo Building executable with PyInstaller...
pyinstaller converter.spec

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build completed successfully!
    echo Executable location: dist\ExcelTimeTrackerConverter-*.exe
) else (
    echo.
    echo Build failed!
    exit /b 1
)
