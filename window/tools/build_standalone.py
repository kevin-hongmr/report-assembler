# -*- coding: utf-8 -*-
"""打包 Windows「单机版」：解压即用，内置运行环境（可直接运行，无需安装 Python/依赖）。

产出：
  _pkg/汇报汇编程序_单机版/
    ├── app/{main.py, core, ui, fonts_bundled, runtime}
    ├── 启动汇编程序.bat
    ├── 使用说明.md
    └── 图标.png
  汇报汇编程序_单机版.zip
"""
import os
import shutil
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # window/
PKG = os.path.join(ROOT, "_pkg")
NAME = "汇报汇编程序_单机版"
DEST = os.path.join(PKG, NAME)
APP = os.path.join(ROOT, "app")

if os.path.isdir(PKG):
    shutil.rmtree(PKG, ignore_errors=True)
os.makedirs(os.path.join(DEST, "app"), exist_ok=True)

# 1) 程序代码
for f in ("main.py", "__init__.py"):
    shutil.copy2(os.path.join(APP, f), os.path.join(DEST, "app", f))
for sub in ("core", "ui", "fonts_bundled"):
    shutil.copytree(os.path.join(APP, sub), os.path.join(DEST, "app", sub),
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
print("代码复制完成")

# 2) 内置运行环境（含 PyQt6 / python-docx / lxml / PyMuPDF / pywin32）
RT_SRC = os.path.join(APP, "runtime")
if not os.path.isdir(RT_SRC):
    raise SystemExit("!! 未找到 app/runtime，请先从便携版解出运行环境")
shutil.copytree(RT_SRC, os.path.join(DEST, "app", "runtime"),
                ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
print("运行环境复制完成")

# 3) 启动脚本（ASCII + CRLF，避免编码/换行问题）
bat = ('@echo off\n'
       'cd /d "%~dp0"\n'
       'start "" "app\\runtime\\pythonw.exe" "app\\main.py"\n')
with open(os.path.join(DEST, "启动汇编程序.bat"), "w", encoding="ascii", newline="\r\n") as f:
    f.write(bat)

# 4) 图标 + 使用说明
icon = os.path.join(ROOT, "图标.png")
if os.path.isfile(icon):
    shutil.copy2(icon, os.path.join(DEST, "图标.png"))
usage = os.path.join(APP, "README.md")
if os.path.isfile(usage):
    shutil.copy2(usage, os.path.join(DEST, "使用说明.md"))

# 5) 运行时自检
import subprocess  # noqa: E402
r = subprocess.run([os.path.join(DEST, "app", "runtime", "python.exe"), "-c",
                    "import docx, lxml, fitz, PyQt6, win32com.client; print('deps ok')"],
                   capture_output=True, text=True)
print("运行时依赖自检:", (r.stdout + r.stderr).strip())

# 6) 压缩
OUT = os.path.abspath(os.path.join(ROOT, "..", NAME + ".zip"))
print("压缩中 ...")
total = 0
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
    for base, _dirs, files in os.walk(DEST):
        for fn in files:
            full = os.path.join(base, fn)
            arc = os.path.relpath(full, PKG)
            z.write(full, arc)
            total += 1
print("完成：%s（%d 个文件，%.1f MB）" % (OUT, total, os.path.getsize(OUT) / 1048576))
