@echo off
cd /d "%~dp0"
title DocAssembler Server (LAN)

rem Launch order: bundled runtime -> dev venv -> system python
if exist "app\runtime\python.exe" (
    "app\runtime\python.exe" "app\server.py" --port 8000
    goto end
)
if exist "app\.venv\Scripts\python.exe" (
    "app\.venv\Scripts\python.exe" "app\server.py" --port 8000
    goto end
)
where python.exe >nul 2>nul
if not errorlevel 1 (
    python.exe "app\server.py" --port 8000
    goto end
)

echo.
echo  [Python runtime not found]
echo  Install Python 3.8+ and run: pip install -r requirements.txt
echo.
pause
:end
