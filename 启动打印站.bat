@echo off
cd /d "%~dp0"
setlocal

echo ============================================
echo   Print Station - Starting...
echo ============================================

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python was not found on this computer.
    echo [ERROR] Install Python 3.9 or later from https://www.python.org/downloads/
    echo [ERROR] and enable "Add python.exe to PATH", then start this file again.
    pause
    exit /b 1
)

rem --- Avoid starting a second web server if one is already running ---
netstat -ano | findstr ":5000" | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 (
    echo [INFO] Web server already running on port 5000.
) else (
    start "PrintStation-Web" cmd /k python run.py
    ping -n 4 127.0.0.1 >nul
)

set "CPOLAR_EXE=%LOCALAPPDATA%\cpolar\cpolar\cpolar.exe"
if not exist "%CPOLAR_EXE%" if exist "%~dp0cpolar\cpolar.exe" set "CPOLAR_EXE=%~dp0cpolar\cpolar.exe"

set "CPOLAR_CONF=%USERPROFILE%\.cpolar\cpolar.yml"
if not exist "%CPOLAR_CONF%" if exist "%~dp0cpolar\cpolar.yml" set "CPOLAR_CONF=%~dp0cpolar\cpolar.yml"

if not exist "%CPOLAR_EXE%" goto :localonly
if not exist "%CPOLAR_CONF%" goto :localonly
findstr /i /c:"authtoken:" "%CPOLAR_CONF%" >nul 2>&1
if errorlevel 1 goto :localonly

rem --- Avoid starting a second cpolar tunnel if one is already running ---
tasklist /fi "imagename eq cpolar.exe" 2>nul | findstr /i "cpolar.exe" >nul
if not errorlevel 1 (
    echo [INFO] cpolar tunnel already running.
    if exist "%~dp0cpolar-subdomain.txt" echo [INFO] Site should be available at: https://print123.cpolar.cn
    ping -n 6 127.0.0.1 >nul
    exit /b 0
)

set "SUBDOMAIN="
if exist "%~dp0cpolar-subdomain.txt" set /p SUBDOMAIN=<"%~dp0cpolar-subdomain.txt"

if defined SUBDOMAIN (
    "%CPOLAR_EXE%" http 5000 -log stdout -region cn -subdomain %SUBDOMAIN%
) else (
    "%CPOLAR_EXE%" http 5000 -log stdout -region cn
)
pause
exit /b 0

:localonly
echo [INFO] cpolar is not configured, so public access is disabled.
echo [INFO] Local site: http://127.0.0.1:5000
echo [INFO] To enable public access, install cpolar, sign in, and put the
echo [INFO] authtoken in %%USERPROFILE%%\.cpolar\cpolar.yml, then start again.
echo [INFO] To use a reserved name, put it in cpolar-subdomain.txt.
ping -n 6 127.0.0.1 >nul
exit /b 0
