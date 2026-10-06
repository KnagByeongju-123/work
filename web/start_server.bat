@echo off
cd /d "%~dp0"
title TJD MES Web Server
set PYTHONIOENCODING=utf-8
set PORT=8000
set PY=python
where python >nul 2>nul || set PY=py
rem Open the web page 3 seconds later (server starts first)
start "" /min cmd /c "timeout /t 3 /nobreak >nul & start http://localhost:%PORT%"
%PY% server.py %PORT%
rem Normal stop (all web pages closed) = close this window. Error = keep it open.
if %errorlevel% equ 0 exit
echo.
echo [Server stopped] Check the message above.
pause
