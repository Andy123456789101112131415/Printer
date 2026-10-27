@echo off
cd /d "%~dp0"

echo ============================================
echo   Print Station - Starting...
echo   Public URL: https://print123.cpolar.cn
echo ============================================

rem --- Avoid starting a second web server if one is already running ---
netstat -ano | findstr ":5000" | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 (
    echo [INFO] Web server already running on port 5000.
) else (
    start "PrintStation-Web" cmd /k python run.py
    ping -n 4 127.0.0.1 >nul
)

rem --- Avoid starting a second cpolar tunnel if one is already running ---
tasklist /fi "imagename eq cpolar.exe" 2>nul | findstr /i "cpolar.exe" >nul
if not errorlevel 1 (
    echo [INFO] cpolar tunnel already running.
    echo [INFO] Site should be available at: https://print123.cpolar.cn
    ping -n 6 127.0.0.1 >nul
    exit /b 0
)

"%LOCALAPPDATA%\cpolar\cpolar\cpolar.exe" http 5000 -log stdout -region cn -subdomain print123

pause
