@echo off
cd /d "%~dp0"

echo ============================================
echo   Print Station - Starting...
echo   Window 1 = web server (auto-restart)
echo   Window 2 = cpolar tunnel (fixed URL)
echo   Public URL: https://print123.cpolar.cn
echo ============================================

start "PrintStation-Web" cmd /k python run.py

timeout /t 2 >nul

"%LOCALAPPDATA%\cpolar\cpolar\cpolar.exe" http 5000 -log stdout -region cn -subdomain print123

pause
