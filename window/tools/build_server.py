# -*- coding: utf-8 -*-
"""打包 Windows「服务器版」：解压即用，局域网内浏览器访问。

产出：
  _pkg/汇报汇编程序_服务器版/
    ├── app/{__init__.py, core, server.py, web, fonts_bundled, runtime}
    ├── 启动服务器版.bat
    ├── 使用说明_服务器版.md
    └── 图标.png
  汇报汇编程序_服务器版.zip
"""
import os
import shutil
import subprocess
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # window/
PKG = os.path.join(ROOT, "_pkg_server")
NAME = "汇报汇编程序_服务器版"
DEST = os.path.join(PKG, NAME)
APP = os.path.join(ROOT, "app")

if os.path.isdir(PKG):
    shutil.rmtree(PKG, ignore_errors=True)
os.makedirs(os.path.join(DEST, "app"), exist_ok=True)

# 1) 程序代码（服务器版不需要 ui/main.py）
shutil.copy2(os.path.join(APP, "__init__.py"), os.path.join(DEST, "app", "__init__.py"))
shutil.copy2(os.path.join(APP, "server.py"), os.path.join(DEST, "app", "server.py"))
for sub in ("core", "web", "fonts_bundled"):
    shutil.copytree(os.path.join(APP, sub), os.path.join(DEST, "app", sub),
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
print("代码复制完成")

# 2) 内置运行环境
RT_SRC = os.path.join(APP, "runtime")
if not os.path.isdir(RT_SRC):
    raise SystemExit("!! 未找到 app/runtime")
shutil.copytree(RT_SRC, os.path.join(DEST, "app", "runtime"),
                ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
print("运行环境复制完成")

# 3) 启动脚本（ASCII + CRLF；pythonw 无控制台窗口，右下角出现托盘图标）
bat = ('@echo off\n'
       'cd /d "%~dp0"\n'
       'start "" "app\\runtime\\pythonw.exe" "app\\server.py" --port 8000\n')
with open(os.path.join(DEST, "启动服务器版.bat"), "w", encoding="ascii", newline="\r\n") as f:
    f.write(bat)

# 4) 说明与图标
_usage = os.path.join(ROOT, "使用说明_服务器版.md")
if os.path.isfile(_usage):
    shutil.copy2(_usage, os.path.join(DEST, "使用说明_服务器版.md"))
_icon = os.path.join(ROOT, "图标.png")
if os.path.isfile(_icon):
    shutil.copy2(_icon, os.path.join(DEST, "图标.png"))

# 5) 运行时自检
r = subprocess.run([os.path.join(DEST, "app", "runtime", "python.exe"), "-c",
                    "import docx, lxml, fitz, win32com.client; print('deps ok')"],
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
            z.write(full, os.path.relpath(full, PKG))
            total += 1
print("完成：%s（%d 个文件，%.1f MB）" % (OUT, total, os.path.getsize(OUT) / 1048576))
