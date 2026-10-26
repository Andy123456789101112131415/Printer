@echo off
cd /d "%~dp0"

echo ============================================
echo   Print Station - Setup
echo ============================================

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo.
    echo Please install Python 3.9+ from https://www.python.org/downloads/
    echo During install, make sure to CHECK "Add python.exe to PATH".
    echo After installing, run this setup again.
    echo.
    pause
    exit /b 1
)

python --version
echo Installing dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies. Check internet connection and retry.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Setup done!
echo   Next: double-click 启动打印站.bat to start.
echo ============================================
pause
