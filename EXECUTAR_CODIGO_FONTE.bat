@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py -m pip install -r requirements.txt
  py main.py
) else (
  python -m pip install -r requirements.txt
  python main.py
)
if errorlevel 1 pause
