@echo off
REM Run server so you can open it from your mobile on the same Wi-Fi.
REM On phone browser use: http://192.168.1.11:8000
echo Starting server - from phone open http://192.168.1.11:8000
python manage.py runserver 0.0.0.0:8000
pause
