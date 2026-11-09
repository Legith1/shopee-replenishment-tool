@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo Checking Python environment...
set "PY_CMD="

if exist "D:\Program Files\Python313\python.exe" (
    "D:\Program Files\Python313\python.exe" -c "import openpyxl" >nul 2>&1
    if not errorlevel 1 set "PY_CMD=D:\Program Files\Python313\python.exe"
)

if not defined PY_CMD (
    python -c "import openpyxl" >nul 2>&1
    if not errorlevel 1 set "PY_CMD=python"
)

if not defined PY_CMD (
    echo ERROR: Python with openpyxl was not found.
    echo Install openpyxl with: python -m pip install openpyxl
    pause
    exit /b 1
)

echo Python: %PY_CMD%
"%PY_CMD%" --version
echo Starting local server. Close this window to stop it.
"%PY_CMD%" server.py
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" echo ERROR: server stopped with code %EXIT_CODE%
pause
exit /b %EXIT_CODE%
