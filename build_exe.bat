@echo off
REM Build script for creating .exe file from Python application
REM Usage: build_exe.bat

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

REM Clean previous builds
if exist "dist" (
    echo Cleaning previous builds...
    rmdir /s /q dist
)
if exist "build" (
    rmdir /s /q build
)

REM Build executable using spec file
echo Building executable...
pyinstaller converter.spec

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build completed successfully!
    echo Executable location: dist\ExcelTimeTrackerConverter.exe
) else (
    echo.
    echo Build failed!
    exit /b 1
)
