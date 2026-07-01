@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Python virtual environment was not found.
  echo Run setup first: python -m venv .venv
  pause
  exit /b 1
)

".venv\Scripts\python.exe" "src\refresh_exports.py"
set EXITCODE=%ERRORLEVEL%

echo.
if "%EXITCODE%"=="0" (
  echo Export refresh finished.
) else (
  echo Export refresh finished with errors. Check logs or screen message.
)
pause
exit /b %EXITCODE%
