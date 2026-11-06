@echo off
cd /d "%~dp0"

rem Launch order:
rem   1) bundled runtime  app\runtime  (portable release)
rem   2) dev venv         app\.venv    (created locally)
rem   3) system Python    (requires: pip install -r requirements.txt)

if exist "app\runtime\pythonw.exe" (
    start "" "app\runtime\pythonw.exe" "app\main.py"
    exit /b 0
)

if exist "app\.venv\Scripts\pythonw.exe" (
    start "" "app\.venv\Scripts\pythonw.exe" "app\main.py"
    exit /b 0
)

where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pythonw.exe "app\main.py"
    exit /b 0
)

where python.exe >nul 2>nul
if not errorlevel 1 (
    start "" python.exe "app\main.py"
    exit /b 0
)

echo.
echo  [Python runtime not found]
echo.
echo  Do one of the following:
echo    1. Install Python 3.8+ then run:  pip install -r requirements.txt
echo    2. Use the portable release (runtime is bundled, no install needed)
echo.
pause
exit /b 1
