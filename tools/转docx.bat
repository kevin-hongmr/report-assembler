@echo off
chcp 65001 >nul
rem ============================================
rem  批量把 .doc/.wps 转成 .docx（仅依赖 WPS）
rem  用法：把 .doc/.wps 文件或文件夹直接拖到本图标上，松开即批量转换。
rem  转换后的 .docx 输出在原位置，可直接添加进汇编程序。
rem ============================================
setlocal
set "ROOT=%~dp0.."
set "PY="
if exist "%ROOT%\app\.venv\Scripts\python.exe" set "PY=%ROOT%\app\.venv\Scripts\python.exe"
if not defined PY set "PY=python"

if "%~1"=="" (
  echo 请把 .doc/.wps 文件或文件夹拖到本图标上，松开即批量转换。
  echo 或双击后在命令行输入：python tools\convert_to_docx.py 文件或文件夹 [--outdir 目录]
  echo.
  "%PY%" "%ROOT%\tools\convert_to_docx.py" --help
  echo.
  pause
  exit /b 0
)

"%PY%" "%ROOT%\tools\convert_to_docx.py" %*
echo.
pause
