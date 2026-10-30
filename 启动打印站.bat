@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Print Station

echo ============================================
echo   Print Station - One-click setup and start
echo ============================================

rem ---------- 1. Find or install Python ----------
set "PY="
for /f "delims=" %%p in ('where python 2^>nul') do if not defined PY set "PY=%%p"
if not defined PY for /f "delims=" %%p in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%p"

if defined PY goto :check_packages

echo [SETUP] Python was not found on this computer.
if "%PS_DRYRUN%"=="1" (
    echo [DRYRUN] would download the Python installer and install it silently
    pause
    exit /b 0
)

echo [SETUP] Downloading Python 3.11 ...
set "INSTALLER=%TEMP%\printstation-python-setup.exe"
del /q "%INSTALLER%" 2>nul
curl -L --fail --silent --show-error -o "%INSTALLER%" https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
if not exist "%INSTALLER%" powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -UseBasicParsing -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile '%INSTALLER%' } catch { exit 1 }"

if not exist "%INSTALLER%" (
    echo [ERROR] Python could not be downloaded. Check the internet connection,
    echo [ERROR] or install Python 3.9+ from https://www.python.org/downloads/
    echo [ERROR] with "Add python.exe to PATH" enabled, then run this file again.
    pause
    exit /b 1
)

echo [SETUP] Installing Python, this can take a few minutes ...
"%INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1 Include_test=0
ping -n 20 127.0.0.1 >nul
del /q "%INSTALLER%" 2>nul

set "PY="
for /f "delims=" %%p in ('where python 2^>nul') do if not defined PY set "PY=%%p"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not defined PY (
    echo [ERROR] Python installation did not finish. Install Python 3.9+ manually
    echo [ERROR] with "Add python.exe to PATH" enabled, then run this file again.
    pause
    exit /b 1
)

:check_packages
echo [SETUP] Checking required Python packages ...
"%PY%" -c "import flask, waitress, win32print" >nul 2>&1
if not errorlevel 1 goto :check_printer

echo [SETUP] Installing required packages, this can take a few minutes ...
"%PY%" -m pip install --disable-pip-version-check --quiet -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Could not install the Python packages. Check the internet
    echo [ERROR] connection and run this file again.
    pause
    exit /b 1
)

:check_printer
echo [SETUP] Checking the school printer ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$n='\\172.16.14.98\wccprinter'; if (Get-Printer -Name $n -ErrorAction SilentlyContinue) { exit 0 }; try { Add-Printer -ConnectionName $n -ErrorAction Stop; exit 0 } catch { Write-Host ('[WARN] Could not add the printer: ' + $_.Exception.Message); exit 1 }"

echo [SETUP] Keeping the computer awake while plugged in ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "powercfg /change standby-timeout-ac 0" >nul 2>&1

rem ---------- 2. Start the web server ----------
netstat -ano | findstr ":5000" | findstr "LISTENING" >nul 2>&1
if errorlevel 1 (
    start "PrintStation-Web" cmd /k ""%PY%" run.py"
    ping -n 4 127.0.0.1 >nul
) else (
    echo [INFO] Web server already running on port 5000.
)

rem ---------- 3. Start the public tunnel ----------
set "CPOLAR_EXE=%LOCALAPPDATA%\cpolar\cpolar\cpolar.exe"
if not exist "%CPOLAR_EXE%" if exist "%~dp0cpolar\cpolar.exe" set "CPOLAR_EXE=%~dp0cpolar\cpolar.exe"

set "CPOLAR_CONF=%USERPROFILE%\.cpolar\cpolar.yml"
if not exist "%CPOLAR_CONF%" if exist "%~dp0cpolar\cpolar.yml" set "CPOLAR_CONF=%~dp0cpolar\cpolar.yml"

if not exist "%CPOLAR_EXE%" goto :localonly
if not exist "%CPOLAR_CONF%" goto :localonly
findstr /i /c:"authtoken:" "%CPOLAR_CONF%" >nul 2>&1
if errorlevel 1 goto :localonly

tasklist /fi "imagename eq cpolar.exe" 2>nul | findstr /i "cpolar.exe" >nul
if not errorlevel 1 (
    echo [INFO] Public tunnel already running.
    echo [INFO] Local site:  http://127.0.0.1:5000
    echo [INFO] Public site: https://print123.cpolar.cn
    ping -n 6 127.0.0.1 >nul
    exit /b 0
)

set "SUBDOMAIN="
if exist "%~dp0cpolar-subdomain.txt" set /p SUBDOMAIN=<"%~dp0cpolar-subdomain.txt"

echo [INFO] Local site: http://127.0.0.1:5000
if defined SUBDOMAIN echo [INFO] Public site: https://%SUBDOMAIN%.cpolar.cn
echo.

if defined SUBDOMAIN (
    "%CPOLAR_EXE%" http 5000 -log stdout -region cn -subdomain %SUBDOMAIN% -config "%CPOLAR_CONF%"
) else (
    "%CPOLAR_EXE%" http 5000 -log stdout -region cn -config "%CPOLAR_CONF%"
)

echo.
echo [INFO] The tunnel stopped. If it said the reserved name is already in use,
echo [INFO] stop Print Station on the other computer first, then start this file again.
pause
exit /b 0

:localonly
echo [INFO] Public access is not configured, so only this computer can open the site.
echo [INFO] Local site: http://127.0.0.1:5000
echo [INFO] To publish it, put cpolar.yml with an authtoken in the cpolar folder.
ping -n 6 127.0.0.1 >nul
exit /b 0
