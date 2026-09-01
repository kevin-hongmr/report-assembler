# -*- coding: utf-8 -*-
"""打包脚本：把程序打包成可便携解压运行的目录，再压缩成 zip。

产出：
  _pkg/汇报汇编程序/
    ├── 使用说明.md
    ├── 启动汇编程序.bat
    ├── 启动汇编程序.sh
    └── app/{main.py, core, ui, fonts_bundled, runtime}
  汇报汇编程序_便携版.zip
"""
import os
import shutil
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PKG = os.path.join(ROOT, "_pkg")
DEST = os.path.join(PKG, "汇报汇编程序")
APP_SRC = os.path.join(ROOT, "app")
MANAGED_PY = r"C:\Users\15913\.workbuddy\binaries\python\versions\3.13.12"
VENV_SP = os.path.join(ROOT, "app", ".venv", "Lib", "site-packages")

# 1) 清理并建目录
if os.path.isdir(PKG):
    shutil.rmtree(PKG)
os.makedirs(os.path.join(DEST, "app", "runtime"), exist_ok=True)
os.makedirs(os.path.join(DEST, "app"), exist_ok=True)

# 2) 复制 app 代码（排除 __pycache__ / .venv / libreoffice）
for name in ("main.py", "__init__.py"):
    shutil.copy2(os.path.join(APP_SRC, name), os.path.join(DEST, "app", name))
for sub in ("core", "ui"):
    shutil.copytree(os.path.join(APP_SRC, sub), os.path.join(DEST, "app", sub),
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
shutil.copytree(os.path.join(APP_SRC, "fonts_bundled"),
                os.path.join(DEST, "app", "fonts_bundled"))

# 3) 复制内置 Python 运行时（排除不需要的头文件/导入库/脚本）
RUNTIME = os.path.join(DEST, "app", "runtime")
EXCLUDE_DIRS = {"include", "libs", "Scripts"}
for name in os.listdir(MANAGED_PY):
    if name in EXCLUDE_DIRS:
        continue
    src = os.path.join(MANAGED_PY, name)
    dst = os.path.join(RUNTIME, name)
    if os.path.isdir(src):
        shutil.copytree(src, dst, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    else:
        shutil.copy2(src, dst)

# 4) 把 venv 的 site-packages 依赖合并进 runtime
rt_sp = os.path.join(RUNTIME, "Lib", "site-packages")
os.makedirs(rt_sp, exist_ok=True)
for name in os.listdir(VENV_SP):
    src = os.path.join(VENV_SP, name)
    dst = os.path.join(rt_sp, name)
    if os.path.isdir(src):
        if os.path.isdir(dst):
            shutil.rmtree(dst)
        shutil.copytree(src, dst, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    else:
        shutil.copy2(src, dst)

# 5) 写启动脚本（ASCII-only，避免编码问题）
# 用 pythonw.exe + start 启动，无控制台窗口，bat 启动后立即退出（cmd 窗口自动关闭）
# 注：字符串用 \n，配合 newline="\r\n" 生成干净的 CRLF（避免叠成 \r\r\n）
bat = '@echo off\ncd /d "%~dp0"\nstart "" "app\\runtime\\pythonw.exe" "app\\main.py"\n'
with open(os.path.join(DEST, "启动汇编程序.bat"), "w", encoding="ascii", newline="\r\n") as f:
    f.write(bat)

# Linux 一键启动脚本：直接复制根目录规范版（LF、UTF-8 无 BOM），并设可执行位。
# 内容为依赖自举逻辑（探测 python3 → 建 .venv → pip 安装依赖 → 运行），不引用内置 Windows 运行时。
_sh_src = os.path.join(ROOT, "启动汇编程序.sh")
if os.path.isfile(_sh_src):
    shutil.copy2(_sh_src, os.path.join(DEST, "启动汇编程序.sh"))
    os.chmod(os.path.join(DEST, "启动汇编程序.sh"), 0o755)
else:
    print("警告：未找到根目录 启动汇编程序.sh，将生成旧版占位脚本。")
    sh = '#!/usr/bin/env bash\ncd "$(dirname "$0")"\nexec python3 app/main.py\n'
    with open(os.path.join(DEST, "启动汇编程序.sh"), "w", encoding="utf-8", newline="\n") as f:
        f.write(sh)

# 6) 复制程序图标与使用说明
_icon = os.path.join(ROOT, "图标.png")
if os.path.isfile(_icon):
    shutil.copy2(_icon, os.path.join(DEST, "图标.png"))
shutil.copy2(os.path.join(ROOT, "使用说明.md"), os.path.join(DEST, "使用说明.md"))

print("打包目录构建完成：", DEST)

# 7) 校验运行时
import subprocess
r = subprocess.run([os.path.join(RUNTIME, "python.exe"), "-c",
                    "import docx, lxml, fitz, PyQt6, win32com.client; print('deps ok')"],
                   capture_output=True, text=True)
print("运行时依赖自检：", (r.stdout + r.stderr).strip())
