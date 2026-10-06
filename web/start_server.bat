@echo off
cd /d "%~dp0"
title TJD MES Web Server
set PYTHONIOENCODING=utf-8
set PY=python
where python >nul 2>nul || set PY=py
%PY% server.py %*
echo.
echo [Server stopped] Check the message above.
pause
