# -*- coding: utf-8 -*-
"""就地更新便携版 zip：替换改动的源文件、启动脚本、图标，不重建 100MB 运行时。

## 为什么需要「就地更新」而不是重打包
Windows 便携版内含约 100MB 的 `app/runtime/`（Windows 版 Python）。重新打包要把这一大堆
文件重新拷贝一遍，慢且容易在跨机复制时丢权限。绝大多数改动只涉及少量文本文件，
直接在 zip 内替换对应条目即可，其余条目按原始 `ZipInfo` 原样搬运（保留压缩方式、
时间戳与 UTF-8 文件名标志）。

## 重要：替换清单要区分「跨平台」与「单平台」
本仓库同时产出 Windows 便携版与国产系统离线版。**不是所有文件都能互相搬运**：

* 跨平台共用（可安全替换进任一包）：
  - `app/core/assembler.py` —— OPC 打包与公文排版逻辑，纯 docx/XML，无平台分支
  - `app/core/converter.py` —— 转换引擎调度（Windows 走 WPS COM / Linux 走 LibreOffice
    在同一文件内分支）
  - `app/core/logger.py`、`app/main.py` —— 纯 Python
* 仅国产系统（**不可**搬进 Windows 包）：
  - `app/ui/appicon.py`、`startup` 脚本、`assets/`、`ime/`、`国产系统运行说明.md`
    —— 依赖 xprop / GTK 图标缓存 / fcitx / ELF 架构判断
* 顶层文件按包区分：Windows 用 `启动汇编程序.bat`，国产系统用 `启动汇编程序.sh`（需 0755）。

默认仍是历史的全量清单（当年代码基本共用）。只想补一个文件时用 `--files`，例如
Windows 包只补 OPC 命名空间修复：

    python tools/update_zip.py --src 汇报汇编程序_便携版.zip \\
        --dst 汇报汇编程序_便携版_new.zip --files app/core/assembler.py
"""
import argparse
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOP = "汇报汇编程序"

# zip 内相对路径 -> 本地磁盘上的源文件（全部按“跨平台共用”挑选）
# 说明：main_window.py 内的引擎提示已用 platform.system() 分支，两平台各自显示正确的
# 排障文字，故可安全共用；converter.py 的 Windows 走 WPS/Word COM、Linux 走 LibreOffice
# 也在同一文件内分支。
CROSS_PLATFORM_FILES = [
    "app/main.py",
    "app/ui/main_window.py",
    "app/core/assembler.py",
    "app/core/converter.py",
    "app/core/pipeline.py",
    "app/core/logger.py",
]

# 顶层文件：Linux 启动脚本要设可执行位；.bat 需 CRLF。
# 注意 README.md 不在 Windows 便携版包内（它只随国产系统离线版分发），
# 放进来会凭空多出一个顶层文件，故不列入。
TOP_FILES = ["启动汇编程序.sh", "启动汇编程序.bat", "图标.png"]

# Windows 版启动器（内容固定，不随仓库变化）
WINDOWS_BAT = b'@echo off\r\ncd /d "%~dp0"\r\nstart "" "app\\runtime\\pythonw.exe" "app\\main.py"\r\n'


def build_replacements(selected):
    out = {}
    for rel in selected:
        if rel == "启动汇编程序.bat":
            out["%s/%s" % (TOP, rel)] = WINDOWS_BAT
            continue
        p = os.path.join(ROOT, rel)
        if not os.path.isfile(p):
            print("  [跳过] 工作区缺少 %s" % rel)
            continue
        with open(p, "rb") as f:
            out["%s/%s" % (TOP, rel)] = f.read()
    return out


def make_info(name):
    info = zipfile.ZipInfo(name)
    info.compress_type = zipfile.ZIP_DEFLATED
    if name == "%s/启动汇编程序.sh" % TOP:
        # external_attr 高 16 位为 Unix 权限位：0755
        info.external_attr = 0o100755 << 16
    return info


def main():
    ap = argparse.ArgumentParser(description="就地更新便携版 zip 内的指定文件")
    ap.add_argument("--src", default=os.path.join(ROOT, "汇报汇编程序_便携版.zip"))
    ap.add_argument("--dst", default=os.path.join(ROOT, "汇报汇编程序_便携版_new.zip"))
    ap.add_argument("--files", default=",".join(CROSS_PLATFORM_FILES + TOP_FILES),
                    help="要替换的相对路径，逗号分隔；默认全量清单")
    a = ap.parse_args()

    selected = [x.strip() for x in a.files.split(",") if x.strip()]
    replacements = build_replacements(selected)
    if not replacements:
        print("没有可替换的文件，退出。")
        return 1

    src, dst = os.path.abspath(a.src), os.path.abspath(a.dst)
    if not os.path.isfile(src):
        print("!! 找不到源包 %s" % src)
        return 1
    if os.path.exists(dst):
        print("!! 目标已存在，请先移走：%s" % dst)
        return 1

    print("源包：%s" % src)
    print("目标：%s" % dst)
    print("替换 %d 个文件：" % len(replacements))
    for n in sorted(replacements):
        print("   %-42s %8d 字节" % (n, len(replacements[n])))

    with zipfile.ZipFile(src, "r") as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        replaced = set()
        for info in zin.infolist():
            name = info.filename
            if name in replacements:
                if name == "%s/启动汇编程序.sh" % TOP:
                    info.external_attr = 0o100755 << 16
                zout.writestr(info, replacements[name])
                replaced.add(name)
            else:
                # 原样搬运（保留压缩方式、时间戳、外部属性）
                zout.writestr(info, zin.read(name))
        # 追加旧包中不存在的新条目
        for name, data in sorted(replacements.items()):
            if name not in replaced:
                zout.writestr(make_info(name), data)
                print("新增条目:", name)

    print("\n已更新 %d 个条目，输出：%s" % (len(replaced), dst))
    return 0


if __name__ == "__main__":
    sys.exit(main())
