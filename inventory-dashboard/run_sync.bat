@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Python virtual environment was not found.
  echo Run setup first: python -m venv .venv
  pause
  exit /b 1
)

".venv\Scripts\python.exe" "src\sync_inventory.py"
set EXITCODE=%ERRORLEVEL%

echo.
if "%EXITCODE%"=="0" (
  echo Sync finished.
) else (
  echo Sync finished with errors. Check logs folder for details.
)
pause
exit /b %EXITCODE%
