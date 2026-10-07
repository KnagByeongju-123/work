@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PY=python
where python >nul 2>nul || set PY=py
%PY% pm_check.py %1
