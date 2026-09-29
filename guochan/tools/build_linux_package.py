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
    # Qt6 中文输入相关：随包提供，便于在目标机上自检/自修复（启动脚本会自动调用）
    "tools/check_ime_patch.py",
    "tools/patch_qt6_ime_plugin.py",
    "tools/patch_qt6_inputcontext.py",
    # 图标多尺寸生成（预生成的图标已随包，此工具用于更换图标后重新生成）
    "tools/make_app_icons.py",
]

# 需要整目录一并打进包的资源（相对项目根）
RESOURCE_DIRS = [
    ("assets", "assets"),   # 多尺寸程序图标（任务栏图标依赖，缺了会显示 Python 图标）
    ("ime", "ime"),         # 适配国产系统的 Qt6 fcitx 输入法插件 + 原始副本（供自修复重建）
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


def elf_machine(path):
    """读 ELF 头 e_machine：183 = AArch64，62 = x86-64。失败返回 None。"""
    try:
        with open(path, "rb") as f:
            head = f.read(20)
    except Exception:  # noqa: BLE001
        return None
    if len(head) < 20 or head[:4] != b"\x7fELF":
        return None
    return int.from_bytes(head[18:20], "little")


_ARCH_MACHINE = {"aarch64": 183, "x86_64": 62}


def find_soffice_bin(lo_src):
    """在便携 LibreOffice 树里找 soffice.bin。"""
    for dirpath, _dirnames, filenames in os.walk(lo_src):
        if "soffice.bin" in filenames:
            return os.path.join(dirpath, "soffice.bin")
    return None


def check_engine_arch(lo_src, arch):
    """核验便携引擎架构是否与 --arch 一致。

    这是**易犯且后果严重**的错误：x86_64 包的默认引擎目录是 app/libreoffice（通常放
    aarch64 版），漏传 --libreoffice-dir 就会打出一个「包名写 x86_64、引擎却是 aarch64」
    的废包——装机后表现为无法预览、无法转换，且很难从日志看出原因。
    返回 (是否一致, 提示字符串)。
    """
    want = _ARCH_MACHINE.get(arch)
    so = find_soffice_bin(lo_src)
    if so is None:
        return True, "未找到 soffice.bin，跳过架构核验"
    got = elf_machine(so)
    if got is None:
        return True, "无法读取 soffice.bin 架构，跳过核验"
    name = {183: "aarch64", 62: "x86_64"}.get(got, "未知(%s)" % got)
    if want is not None and got != want:
        return False, ("引擎架构不符：包为 %s，引擎却是 %s（%s）"
                       % (arch, name, os.path.relpath(so, ROOT)))
    return True, "引擎架构 %s，与 --arch %s 一致" % (name, arch)


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

    # 前置检查：便携引擎架构必须与 --arch 一致（漏传 --libreoffice-dir 会打出废包）
    lo_src_chk = os.path.abspath(opts.libreoffice_dir) if opts.libreoffice_dir \
        else os.path.join(ROOT, "app", "libreoffice")
    if os.path.isdir(os.path.join(lo_src_chk, "opt")):
        ok, msg = check_engine_arch(lo_src_chk, arch)
        if not ok:
            print("!! %s" % msg)
            print("   源目录：%s" % lo_src_chk)
            print("   这类包装到目标机器上会「无法预览、无法转换」，请改用对应架构的引擎目录重打：")
            print("     x86_64 : python tools/build_linux_package.py --arch x86_64 \\")
            print("                  --libreoffice-dir build/libreoffice_x86_64/lo")
            print("     aarch64: python tools/build_linux_package.py --arch aarch64")
            print("              （默认用 app/libreoffice，即麒麟源重建的 aarch64 便携引擎）")
            return 1
        print("便携引擎预检：%s" % msg)
    else:
        print("[警告] %s 下没有 opt/，未做引擎预检" % lo_src_chk)

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

        # 3.5) 资源目录：assets/（多尺寸图标）、ime/（适配版输入法插件）
        for rel, arc in RESOURCE_DIRS:
            src = os.path.join(ROOT, rel)
            if not os.path.isdir(src):
                print("  [警告] 缺少资源目录 %s/（相关功能可能不可用）" % rel)
                continue
            before = stats["count"]
            add_tree(zf, src, "%s/%s" % (base, arc), stats)
            print("  + %s/（%d 个文件）" % (rel, stats["count"] - before))
            if rel == "assets":
                need = os.path.join(src, "icons", "256.png")
                if not os.path.isfile(need):
                    print("  [警告] 缺少 assets/icons/256.png，任务栏图标可能仍显示为 Python 图标。")
                    print("        可执行：python tools/make_app_icons.py 重新生成")
            if rel == "ime":
                # 插件按 CPU 架构分目录（ime/<arch>/），旧扁平布局仍兼容
                arch_dirs = [d for d in ("aarch64", "x86_64")
                             if os.path.isdir(os.path.join(src, d))]
                targets = []
                if arch_dirs:
                    print("      随包输入法插件架构：%s" % ", ".join(arch_dirs))
                    for ad in arch_dirs:
                        targets.append((os.path.join(src, ad), "ime/%s/" % ad))
                else:
                    print("      随包输入法插件：扁平布局（建议改为 ime/<arch>/）")
                    targets.append((src, "ime/"))
                for where, label in targets:
                    for need in ("libfcitx-qt6-kylin.so",
                                 "libfcitxplatforminputcontextplugin.so.orig"):
                        if not os.path.isfile(os.path.join(where, need)):
                            print("  [警告] %s 缺少 %s，该架构下 Qt6 中文输入可能无法自动修复。"
                                  % (label, need))

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
