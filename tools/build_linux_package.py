#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""打包"国产系统专用"离线版（银河麒麟 / 统信UOS，aarch64 或 x86_64）。

与 Windows 便携版的差异：
  * 不含 app/runtime（那是 Windows 版 Python，Linux 上不可执行）
  * 含 app/libreoffice（便携 LibreOffice，供 PDF 预览/转换；Windows 版 1.6GB 的
    app/libreoffice 需先替换为 aarch64/x86_64 对应架构的便携版，否则会被跳过）
  * 含 wheels/ 离线依赖包，首次运行即可零联网安装
  * 含 字体包/ 全部字体（含 SimSun.ttf）

用法：
    python tools/build_linux_package.py                # 默认 aarch64
    python tools/build_linux_package.py --arch x86_64
"""

import argparse
import os
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 需要一并打进包的顶层文件
TOP_FILES = [
    "启动汇编程序.sh",
    "图标.png",
    "README.md",
    "国产系统运行说明.md",
    "requirements.txt",
]

# 独立转换脚本（方便用户在国产系统上先把 .doc/.wps 批量转成 .docx，
# 再添加进汇编程序；Windows 上则走 WPS。由 app/core/converter 自动选引擎）
TOOL_FILES = [
    "tools/convert_to_docx.py",
    "tools/转docx.sh",
    "tools/转docx.bat",
]

# 需要排除的目录 / 文件
EXCLUDE_DIRS = {
    "libreoffice",      # 由 add_libreoffice_tree 单独处理（需设可执行位），不在通用遍历中打包
    "runtime",          # Windows 版 Python
    ".venv",            # 本机开发虚拟环境
    "__pycache__",
    ".git",
    ".idea",
    ".workbuddy",
    ".planning",
}
EXCLUDE_SUFFIX = (".pyc", ".pyo", ".zip", ".log", ".tmp")


def sync_fonts():
    """把 字体包/ 下的字体（尤其是 SimSun.ttf）同步进 app/fonts_bundled/。

    fonts.py 会检查宋体(simsun)，麒麟系统常缺该字体，因此一并内置。
    """
    src_dir = os.path.join(ROOT, "字体包")
    dst_dir = os.path.join(ROOT, "app", "fonts_bundled")
    os.makedirs(dst_dir, exist_ok=True)
    copied = []
    if os.path.isdir(src_dir):
        for fn in sorted(os.listdir(src_dir)):
            if not fn.lower().endswith((".ttf", ".ttc", ".otf")):
                continue
            s = os.path.join(src_dir, fn)
            d = os.path.join(dst_dir, fn)
            if not os.path.exists(d) or os.path.getsize(s) != os.path.getsize(d):
                shutil.copy2(s, d)
                copied.append(fn)
    return copied


def add_tree(zf, src_dir, arc_root, stats):
    """把一个目录加入 zip，跳过排除项。"""
    for dirpath, dirnames, filenames in os.walk(src_dir):
        dirnames[:] = [d for d in dirnames if d not in EXCLUDE_DIRS]
        rel_dir = os.path.relpath(dirpath, src_dir)
        for fn in filenames:
            if fn.endswith(EXCLUDE_SUFFIX) or fn.startswith("."):
                continue
            full = os.path.join(dirpath, fn)
            rel = fn if rel_dir == "." else os.path.join(rel_dir, fn)
            arc = "%s/%s" % (arc_root, rel.replace("\\", "/"))
            zf.write(full, arc)
            stats["count"] += 1
            stats["bytes"] += os.path.getsize(full)


def add_libreoffice_tree(zf, src_dir, arc_root, stats):
    """把便携 LibreOffice（app/libreoffice，deb 解出后的 opt/libreoffice*/ 结构）加入 zip。

    program/ 下的文件设为可执行位（soffice / soffice.bin 需可执行才能被调用）。
    说明：部分解压工具会丢失可执行位，故 启动汇编程序.sh 内还会再 chmod 兜底一次。
    """
    if not os.path.isdir(src_dir):
        return
    for dirpath, dirnames, filenames in os.walk(src_dir):
        # usr/ 目录是桌面集成（启动器软链 + .desktop），无头转换用不到，且软链在 Windows 上无法 stat，直接跳过
        dirnames[:] = [d for d in dirnames if d not in EXCLUDE_DIRS and d != "usr"]
        rel_dir = os.path.relpath(dirpath, src_dir).replace("\\", "/")
        in_program = rel_dir.endswith("program") or "/program/" in rel_dir
        for fn in filenames:
            if fn.endswith(EXCLUDE_SUFFIX) or fn.startswith("."):
                continue
            full = os.path.join(dirpath, fn)
            if os.path.islink(full):  # 跳过软链接（usr/ 已排除，此处为兜底）
                continue
            rel = fn if rel_dir == "." else os.path.join(rel_dir, fn)
            arc = "%s/%s" % (arc_root, rel.replace("\\", "/"))
            if in_program:
                info = zipfile.ZipInfo(arc)
                info.external_attr = 0o100755 << 16
                info.compress_type = zipfile.ZIP_DEFLATED
                with open(full, "rb") as f:
                    zf.writestr(info, f.read())
            else:
                zf.write(full, arc)
            stats["count"] += 1
            stats["bytes"] += os.path.getsize(full)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--arch", default="aarch64", choices=["aarch64", "x86_64"])
    ap.add_argument("--out", default=None, help="输出 zip 路径")
    ap.add_argument("--wheels-dir", default=None,
                    help="离线依赖目录，默认 build/wheels/<arch>（即 download_wheels.py 的输出）。"
                         "如需手动指定，可传 build/wheels/aarch64 等。")
    ap.add_argument("--libreoffice-dir", default=None,
                    help="便携 LibreOffice 目录（含 opt/ 子目录），默认 app/libreoffice。"
                         "用于 aarch64 / x86_64 各自解包目录互不覆盖。")
    opts = ap.parse_args()

    arch = opts.arch
    out = opts.out or os.path.join(ROOT, "汇报汇编程序_国产系统专用_%s.zip" % arch)

    print("=" * 70)
    print("打包国产系统专用离线版（%s）" % arch)
    print("=" * 70)

    # 前置检查
    wheels = os.path.abspath(opts.wheels_dir) if opts.wheels_dir else os.path.join(ROOT, "build", "wheels", arch)
    if not os.path.isdir(wheels) or not [f for f in os.listdir(wheels) if f.endswith(".whl")]:
        print("!! wheels/ 为空，请先执行：")
        print("     python tools/download_wheels.py --arch %s" % arch)
        return 1

    sh = os.path.join(ROOT, "启动汇编程序.sh")
    if not os.path.isfile(sh):
        print("!! 缺少 启动汇编程序.sh")
        return 1

    # 同步字体
    copied = sync_fonts()
    if copied:
        print("已同步字体到 app/fonts_bundled/：%s" % ", ".join(copied))

    stats = {"count": 0, "bytes": 0}
    base = "汇报汇编程序"

    if os.path.exists(out):
        os.remove(out)

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        # 1) 顶层文件（.sh 需要可执行位）
        for fn in TOP_FILES:
            p = os.path.join(ROOT, fn)
            if not os.path.isfile(p):
                print("  [跳过] 顶层文件不存在：%s" % fn)
                continue
            if fn.endswith(".sh"):
                info = zipfile.ZipInfo("%s/%s" % (base, fn))
                info.external_attr = 0o100755 << 16
                info.compress_type = zipfile.ZIP_DEFLATED
                with open(p, "rb") as f:
                    zf.writestr(info, f.read())
            else:
                zf.write(p, "%s/%s" % (base, fn))
            stats["count"] += 1
            stats["bytes"] += os.path.getsize(p)
            print("  + %s" % fn)

        # 1.5) 独立转换脚本（.sh 需要可执行位）
        for fn in TOOL_FILES:
            p = os.path.join(ROOT, fn)
            if not os.path.isfile(p):
                print("  [跳过] 工具脚本不存在：%s" % fn)
                continue
            arc = "%s/%s" % (base, fn)
            if fn.endswith((".sh", ".bat")):
                info = zipfile.ZipInfo(arc)
                info.external_attr = (0o100755 if fn.endswith(".sh") else 0o100644) << 16
                info.compress_type = zipfile.ZIP_DEFLATED
                with open(p, "rb") as f:
                    zf.writestr(info, f.read())
            else:
                zf.write(p, arc)
            stats["count"] += 1
            stats["bytes"] += os.path.getsize(p)
            print("  + %s" % fn)

        # 2) app/（排除 runtime / libreoffice / .venv）
        app_src = os.path.join(ROOT, "app")
        before = stats["count"]
        add_tree(zf, app_src, base + "/app", stats)
        print("  + app/（%d 个文件）" % (stats["count"] - before))

        # 2.5) 便携 LibreOffice（国产系统 PDF 预览/转换引擎，免 root）
        lo_src = opts.libreoffice_dir or os.path.join(ROOT, "app", "libreoffice")
        if os.path.isdir(os.path.join(lo_src, "opt")):
            before = stats["count"]
            add_libreoffice_tree(zf, lo_src, base + "/app/libreoffice", stats)
            print("  + app/libreoffice（便携 LibreOffice，%d 个文件）" %
                  (stats["count"] - before))
        else:
            print("  [警告] 未找到 app/libreoffice/opt（便携 LibreOffice），"
                  "国产系统 PDF 预览将不可用")

        # 3) 字体包/
        font_src = os.path.join(ROOT, "字体包")
        if os.path.isdir(font_src):
            before = stats["count"]
            add_tree(zf, font_src, base + "/字体包", stats)
            print("  + 字体包/（%d 个文件）" % (stats["count"] - before))

        # 4) wheels/
        before = stats["count"]
        nw = 0
        for fn in sorted(os.listdir(wheels)):
            if not fn.endswith(".whl"):
                continue
            zf.write(os.path.join(wheels, fn), "%s/wheels/%s" % (base, fn))
            nw += 1
            stats["count"] += 1
            stats["bytes"] += os.path.getsize(os.path.join(wheels, fn))
        print("  + wheels/（%d 个 wheel）" % nw)

    size_mb = os.path.getsize(out) / 1048576.0
    print("\n" + "=" * 70)
    print("输出：%s" % out)
    print("文件数：%d    源体积：%.1f MB    压缩包：%.1f MB" %
          (stats["count"], stats["bytes"] / 1048576.0, size_mb))
    print("=" * 70)
    print("\n请确认压缩包内架构与 wheel 一致：")
    print("    python tools/verify_wheels.py --arch %s --wheels %s" % (arch, wheels))
    return 0


if __name__ == "__main__":
    sys.exit(main())
