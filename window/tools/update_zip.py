# -*- coding: utf-8 -*-
"""就地更新便携版 zip：替换改动的源文件、启动脚本、图标，追加 logger.py。

运行后生成 *_new.zip，随后手动覆盖原名（或直接用 new 名交付）。
不重建 100MB 运行时，仅更新少量文本文件与图标，避免大目录 rmtree 与重复拷贝。
"""
import os
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_ZIP = os.path.join(ROOT, "汇报汇编程序_便携版.zip")
DST_ZIP = os.path.join(ROOT, "汇报汇编程序_便携版_new.zip")
TOP = "汇报汇编程序"

# 新增/替换的文件（zip 内相对路径 -> 本地磁盘绝对路径或字节）
new_bat = b'@echo off\r\ncd /d "%~dp0"\r\nstart "" "app\\runtime\\pythonw.exe" "app\\main.py"\r\n'
# Linux 启动脚本：读取根目录规范版（LF、UTF-8 无 BOM），杜绝 CRLF 导致的 ^M 解释器错误
new_sh = open(os.path.join(ROOT, "启动汇编程序.sh"), "rb").read()
# 程序图标
new_icon = open(os.path.join(ROOT, "图标.png"), "rb").read()

replacements = {
    f"{TOP}/启动汇编程序.bat": new_bat,
    f"{TOP}/启动汇编程序.sh": new_sh,
    f"{TOP}/图标.png": new_icon,
    f"{TOP}/README.md": open(os.path.join(ROOT, "README.md"), "rb").read(),
    f"{TOP}/app/main.py": open(os.path.join(ROOT, "app", "main.py"), "rb").read(),
    f"{TOP}/app/ui/main_window.py": open(os.path.join(ROOT, "app", "ui", "main_window.py"), "rb").read(),
    f"{TOP}/app/core/assembler.py": open(os.path.join(ROOT, "app", "core", "assembler.py"), "rb").read(),
    f"{TOP}/app/core/pipeline.py": open(os.path.join(ROOT, "app", "core", "pipeline.py"), "rb").read(),
    f"{TOP}/app/core/logger.py": open(os.path.join(ROOT, "app", "core", "logger.py"), "rb").read(),
}


def _make_info(name):
    """构造 zip 条目，Linux 启动脚本带可执行位。"""
    info = zipfile.ZipInfo(name)
    info.compress_type = zipfile.ZIP_DEFLATED
    if name == f"{TOP}/启动汇编程序.sh":
        # external_attr 高 16 位为 Unix 权限：0755
        info.external_attr = 0o100755 << 16
    return info


with zipfile.ZipFile(SRC_ZIP, "r") as zin, \
     zipfile.ZipFile(DST_ZIP, "w", zipfile.ZIP_DEFLATED) as zout:
    replaced = set()
    for info in zin.infolist():
        name = info.filename
        if name in replacements:
            if name == f"{TOP}/启动汇编程序.sh":
                info.external_attr = 0o100755 << 16
            zout.writestr(info, replacements[name])
            replaced.add(name)
        else:
            # 原样复制（含压缩方式、时间戳、UTF-8 文件名标志）
            zout.writestr(info, zin.read(name))
    # 追加任何未在旧包中出现的新文件（如 logger.py / 图标 / .sh 若原本不存在）
    for name, data in replacements.items():
        if name not in replaced:
            zout.writestr(_make_info(name), data)
            print("新增条目:", name)

print("已更新条目:")
for n in sorted(replaced):
    print("  ", n)
print("输出:", DST_ZIP)
