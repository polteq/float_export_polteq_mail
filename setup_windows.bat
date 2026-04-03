@echo off
setlocal
cd /d "%~dp0"

echo =========================================
echo Polteq Timesheet Processor Setup (Windows)
echo =========================================

echo Checking for Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH. Please install Python 3.
    pause
    exit /b 1
)

echo Creating virtual environment...
if not exist "venv" (
    python -m venv venv
)

echo Activating virtual environment and installing dependencies...
call venv\Scripts\activate
pip install -r requirements.txt

echo.
echo Running initial configuration...
python shareable_processor.py --setup

echo.
echo Creating Desktop Drag-and-Drop Shortcut...
set SHORTCUT_PATH="%USERPROFILE%\Desktop\Process Timesheet.bat"

echo @echo off > %SHORTCUT_PATH%
echo setlocal >> %SHORTCUT_PATH%
echo cd /d "%~dp0" >> %SHORTCUT_PATH%
echo call venv\Scripts\activate >> %SHORTCUT_PATH%
echo python shareable_processor.py "%%~1" >> %SHORTCUT_PATH%
echo pause >> %SHORTCUT_PATH%

echo.
echo Setup Complete!
echo You can now drag and drop your Timesheet CSV files onto the "Process Timesheet.bat" shortcut on your desktop.
pause
