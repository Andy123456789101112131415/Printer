@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Print Station - Stop

echo ============================================
echo   Print Station - Stopping
echo ============================================

powershell -NoProfile -ExecutionPolicy Bypass -Command "$self='%~dp0'.TrimEnd('\'); $c = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1; if ($c) { $child = [int]$c.OwningProcess; $p = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $child) -ErrorAction SilentlyContinue; if ($p -and $p.ParentProcessId) { Stop-Process -Id $p.ParentProcessId -Force -ErrorAction SilentlyContinue }; Stop-Process -Id $child -Force -ErrorAction SilentlyContinue; Write-Host '[INFO] Web server stopped.' } else { Write-Host '[INFO] Web server was not running.' }; $tunnels = Get-CimInstance Win32_Process -Filter \"Name='cpolar.exe'\" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like ('*' + $self + '*') }; if (-not $tunnels) { $tunnels = Get-CimInstance Win32_Process -Filter \"Name='cpolar.exe'\" -ErrorAction SilentlyContinue }; foreach ($t in $tunnels) { Stop-Process -Id $t.ProcessId -Force -ErrorAction SilentlyContinue }; if ($tunnels) { Write-Host '[INFO] Public tunnel stopped.' } else { Write-Host '[INFO] Public tunnel was not running.' }"

echo [INFO] Print Station is stopped. Uploads and settings are kept.
ping -n 5 127.0.0.1 >nul
exit /b 0
