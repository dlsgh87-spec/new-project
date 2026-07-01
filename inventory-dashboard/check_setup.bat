@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Python virtual environment was not found.
  echo Run setup first: python -m venv .venv
  pause
  exit /b 1
)

".venv\Scripts\python.exe" "src\check_setup.py"
set EXITCODE=%ERRORLEVEL%

echo.
pause
exit /b %EXITCODE%
