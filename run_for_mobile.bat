@echo off
REM Run server so you can open it from your mobile (same Wi-Fi or hotspot).
REM 1. Connect: PC and phone on same network (phone hotspot OR same Wi-Fi)
REM 2. Run this script
REM 3. On phone browser, open the URL shown below

echo.
echo ========================================
echo   LJIET Attendance - Mobile Access
echo ========================================
echo.

REM Get this PC's IP (works with hotspot, Wi-Fi, etc.)
for /f "delims=" %%i in ('powershell -Command "(Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notmatch '^127\.' } | Select-Object -First 1).IPAddress" 2^>nul') do set "LOCAL_IP=%%i"
if "%LOCAL_IP%"=="" (
    echo Could not detect IP. Try: http://YOUR_PC_IP:8000
    echo Find your IP: ipconfig ^| findstr IPv4
    echo.
) else (
    echo On your PHONE browser open:
    echo.
    echo   http://%LOCAL_IP%:8000
    echo.
    echo (Phone and PC must be on same network - hotspot or Wi-Fi)
    echo.
)

echo Starting server...
echo Press Ctrl+C to stop.
echo.
python manage.py runserver 0.0.0.0:8000
pause
