@echo off
setlocal
cd /d "%~dp0"

set "MESSAGE=%~1"
if "%MESSAGE%"=="" set "MESSAGE=Update project"

"C:\Program Files\Git\bin\git.exe" pull --ff-only
if errorlevel 1 goto failed

"C:\Program Files\Git\bin\git.exe" add .
"C:\Program Files\Git\bin\git.exe" diff --cached --quiet
if %errorlevel%==0 (
  echo No changes to save.
  "C:\Program Files\Git\bin\git.exe" status --short --branch
  pause
  exit /b 0
)

"C:\Program Files\Git\bin\git.exe" commit -m "%MESSAGE%"
if errorlevel 1 goto failed

"C:\Program Files\Git\bin\git.exe" push
if errorlevel 1 goto failed

"C:\Program Files\Git\bin\git.exe" status --short --branch
pause
exit /b 0

:failed
echo.
echo Sync failed. Read the message above, then try again.
pause
exit /b 1
