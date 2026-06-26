@echo off
setlocal
cd /d "%~dp0"
"C:\Program Files\Git\bin\git.exe" pull --ff-only
"C:\Program Files\Git\bin\git.exe" status --short --branch
pause
