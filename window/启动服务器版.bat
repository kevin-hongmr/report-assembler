@echo off
cd /d "%~dp0"

rem Launch the server WITHOUT a console window (pythonw + start).
rem A tray icon appears at the bottom-right of the screen.

if exist "app\runtime\pythonw.exe" (
    start "" "app\runtime\pythonw.exe" "app\server.py" --port 12362
    goto end
)
if exist "app\.venv\Scripts\pythonw.exe" (
    start "" "app\.venv\Scripts\pythonw.exe" "app\server.py" --port 12362
    goto end
)
where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pythonw.exe "app\server.py" --port 12362
    goto end
)

echo.
echo  [Python runtime not found]
echo  Install Python 3.8+ and run: pip install -r requirements.txt
echo.
pause
:end
