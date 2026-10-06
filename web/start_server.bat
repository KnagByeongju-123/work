@echo off
chcp 65001 >nul
cd /d "%~dp0"
title TJD MES Web Server
set PYTHONIOENCODING=utf-8
where python >nul 2>nul
if %errorlevel%==0 (
  python server.py %*
) else (
  py server.py %*
)
echo.
echo 서버가 멈췄습니다. 위 메시지를 확인하세요.
pause
