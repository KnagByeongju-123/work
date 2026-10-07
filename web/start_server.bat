@echo off
cd /d "%~dp0"
title TJD MES Web Server
set PYTHONIOENCODING=utf-8
set PORT=8000
set PY=python
where python >nul 2>nul || set PY=py

rem ---- Find the web server .py in this folder (any file name is OK) ----
set SRV=
for %%f in (*.py) do (
  findstr /m /c:"def quit_watch" "%%f" >nul 2>nul && set "SRV=%%f"
)
if not defined SRV (
  echo.
  echo [ERROR] Web server .py file not found in this folder:
  echo         %~dp0
  echo         Put the web server .py, the web page .html and the MES program .py together.
  echo.
  pause
  exit /b 1
)
echo Server file: %SRV%

rem --open : the server opens the web page itself when it is ready
%PY% "%SRV%" %PORT% --open
rem Normal stop (all web pages closed) = close this window. Error = keep it open.
if %errorlevel% equ 0 exit
echo.
echo [Server stopped] Check the message above.
pause
