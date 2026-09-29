#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Qt6 中文输入补丁的「自检 + 自修复」驱动。

## 背景

国产系统（银河麒麟 / 统信 UOS）上，PyQt6 自带的是 Qt 6.7.x，而系统自带的
Qt6 输入法插件（fcitx-frontend-qt6）是按 Qt 6.4 编译的，直接使用会在
事件循环首轮被 Qt 6.7 覆写插件的 `m_watcher` 成员，表现为：

* 界面里**无法输入中文、无法切换输入法**（输入上下文建不起来）；
* 关闭程序时**进程崩溃**（SIGSEGV，退出码 139）。

修法是两处极小的二进制改写，分别由下面两个工具完成：

    tools/patch_qt6_ime_plugin.py     # 改写输入法插件本身（符号兼容）
    tools/patch_qt6_inputcontext.py   # 改写 libQt6Gui 的一条指令（停止覆写）

本脚本把这两步串起来，并提供**幂等**的检查与修复，供启动脚本每次启动前调用：

    python3 tools/check_ime_patch.py --pyqt6 <PyQt6包目录> [--root <项目根>] [--repair]

为什么需要「自修复」：离线包采用「wheels/ 离线安装到 pylibs/」的方案，
一旦 pip 重新安装 PyQt6，`libQt6Gui.so.6` 会被还原成未打补丁的版本，
崩溃与中文输入问题会**复发**。启动时自动补一次即可彻底免疫。

退出码：0 = 状态正常或已修复；1 = 存在无法自动解决的问题（会打印处理建议）。
"""

import argparse
import os
import shutil
import subprocess
import sys

# 随包插件的文件名（相对 ime/ 目录）。
# 2026-09-28 起按 CPU 架构分目录存放：ime/<arch>/<name>，
# 同时保留旧的扁平布局 ime/<name> 作为回退，便于旧包 / 旧部署目录继续可用。
BUNDLED_PLUGIN_NAME = "libfcitx-qt6-kylin.so"
BUNDLED_PLUGIN_ORIG_NAME = "libfcitxplatforminputcontextplugin.so.orig"

PLUGIN_REL = os.path.join("Qt6", "plugins", "platforminputcontexts",
                          "libfcitxplatforminputcontextplugin.so")
GUI_LIB_REL = os.path.join("Qt6", "lib", "libQt6Gui.so.6")

# 随包插件按 Qt 6.7 的头文件编译，只在同一 minor 版本上保证可用
TARGET_QT_PREFIX = "6.7"


def _run(cmd):
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        return p.returncode, (p.stdout or "") + (p.stderr or "")
    except Exception as e:  # noqa: BLE001
        return 1, "执行失败：%s" % e


def qt_version(gui_dir):
    """从 libQt6Core.so.6 中读出 Qt 版本号（不依赖能否 import PyQt6）。"""
    core = os.path.join(os.path.dirname(gui_dir), "libQt6Core.so.6")
    if not os.path.isfile(core):
        # 兼容 lib 目录传进来的情况
        core = os.path.join(gui_dir, "libQt6Core.so.6")
    if not os.path.isfile(core):
        return None
    try:
        with open(core, "rb") as f:
            blob = f.read()
    except Exception:  # noqa: BLE001
        return None
    import re
    # 形如 "Qt 6.7.3 (arm64-little_endian-lp64 shared ...)"
    m = re.search(rb"Qt (\d+\.\d+\.\d+) \([a-z0-9_]+-little_endian", blob)
    if m:
        return m.group(1).decode()
    m = re.search(rb"Qt (\d+\.\d+\.\d+)", blob)
    return m.group(1).decode() if m else None


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


_HOST_MACHINE = {"aarch64": 183, "arm64": 183, "x86_64": 62, "amd64": 62}

# 目录名归一化：不同发行版对同一架构的写法不一（arm64/amd64）
_ARCH_DIR = {"aarch64": "aarch64", "arm64": "aarch64",
             "x86_64": "x86_64", "amd64": "x86_64",
             "i386": "i386", "i686": "i386"}


def host_machine():
    import platform
    m = platform.machine().lower()
    return _HOST_MACHINE.get(m)


def arch_dir():
    """本机架构对应的 ime/ 子目录名。"""
    import platform
    m = platform.machine().lower()
    return _ARCH_DIR.get(m, m)


def resolve_bundled(root, name):
    """按本机架构解析随包文件：优先 ime/<arch>/<name>，回退旧的 ime/<name>。"""
    p = os.path.join(root, "ime", arch_dir(), name)
    if os.path.isfile(p):
        return p
    return os.path.join(root, "ime", name)


def arch_compatible(path):
    """随包二进制是否与本机同架构。判不出来时返回 True（不阻断，交给实际加载去裁定）。"""
    want = host_machine()
    if want is None:
        return True
    got = elf_machine(path)
    return True if got is None else (got == want)


def _tool(root, name):
    p = os.path.join(root, "tools", name)
    return p if os.path.isfile(p) else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pyqt6", required=True, help="PyQt6 包目录（含 Qt6/ 子目录）")
    ap.add_argument("--root", default=os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    help="项目根目录（含 tools/ 与 ime/）")
    ap.add_argument("--repair", action="store_true", help="发现问题时自动修复")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    pkg = os.path.abspath(a.pyqt6)
    plugin = os.path.join(pkg, PLUGIN_REL)
    gui_lib = os.path.join(pkg, GUI_LIB_REL)

    print("=" * 62)
    print("Qt6 中文输入补丁自检")
    print("  PyQt6 : %s" % pkg)
    print("=" * 62)

    if not os.path.isdir(os.path.join(pkg, "Qt6")):
        print("跳过：%s 下没有 Qt6/ 目录（非 PyQt6 布局）。" % pkg)
        return 0

    ver = qt_version(gui_lib)
    print("Qt 版本：%s" % (ver or "未知"))
    if ver and not ver.startswith(TARGET_QT_PREFIX):
        print("提示：随包插件是按 Qt %s 适配的，当前 Qt 为 %s。" % (TARGET_QT_PREFIX, ver))
        print("      若系统自带同版本的 Qt6 输入法插件，将优先使用系统插件；")
        print("      随包插件仅作备选（可能因 Qt 私有符号差异而无法加载）。")

    problems = []
    setup = True

    # ---------- 1) 输入法插件 ----------
    bundled = resolve_bundled(root, BUNDLED_PLUGIN_NAME)
    bundled_orig = resolve_bundled(root, BUNDLED_PLUGIN_ORIG_NAME)
    tool_plugin = _tool(root, "patch_qt6_ime_plugin.py")

    have_bundled = os.path.isfile(bundled) or os.path.isfile(bundled_orig)
    if have_bundled:
        src = bundled if os.path.isfile(bundled) else bundled_orig
        print("随包插件：ime/%s/（%s）" % (arch_dir(), os.path.basename(src)))
    else:
        print("随包插件：未找到本机架构（%s）的随包插件" % arch_dir())

    use_bundled = bool(ver and ver.startswith(TARGET_QT_PREFIX))
    if use_bundled and have_bundled:
        # 跨架构的插件装上去只会让 Qt 加载失败，还会把可用的系统插件挤掉
        probe = bundled if os.path.isfile(bundled) else bundled_orig
        if not arch_compatible(probe):
            got = elf_machine(probe)
            print("随包插件架构与本机不符（插件 e_machine=%s，本机=%s），跳过随包插件，改用系统插件。"
                  % (got, host_machine()))
            use_bundled = False
    if not os.path.isfile(plugin):
        print("输入法插件：缺失")
        if use_bundled and have_bundled:
            if a.repair:
                os.makedirs(os.path.dirname(plugin), exist_ok=True)
                src = bundled if os.path.isfile(bundled) else bundled_orig
                shutil.copy2(src, plugin)
                print("  → 已从随包插件安装：%s" % os.path.basename(src))
                if src == bundled_orig and tool_plugin:
                    rc, out = _run([sys.executable, tool_plugin, bundled_orig, plugin])
                    print("  → 已应用插件补丁（rc=%d）" % rc)
            else:
                problems.append("输入法插件缺失（可自动安装随包插件）")
        else:
            setup = False
            print("  说明：需要使用系统 Qt6 输入法插件，请执行（二选一）：")
            print("        sudo apt install fcitx-frontend-qt6      # fcitx4")
            print("        sudo apt install fcitx5-frontend-qt6     # fcitx5")

    if os.path.isfile(plugin):
        print("输入法插件：%s（%d 字节）" % (os.path.basename(plugin), os.path.getsize(plugin)))
        if tool_plugin:
            rc, out = _run([sys.executable, tool_plugin, "--check", plugin])
            ok = ("RESULT: OK" in out)
            print("  插件补丁：%s" % ("已就绪" if ok else "未就绪"))
            if not ok:
                fixed = False
                if a.repair and use_bundled:
                    # 优先用随包的干净副本重新适配；无则直接用随包已适配副本覆盖
                    if os.path.isfile(bundled_orig):
                        shutil.copy2(bundled_orig, plugin)
                        rc2, out2 = _run([sys.executable, tool_plugin, bundled_orig, plugin])
                        fixed = (rc2 == 0)
                    elif os.path.isfile(bundled):
                        shutil.copy2(bundled, plugin)
                        fixed = True
                    if fixed:
                        rc3, out3 = _run([sys.executable, tool_plugin, "--check", plugin])
                        print("  → 已修复，复核：%s" % ("OK" if "RESULT: OK" in out3 else "仍异常"))
                    else:
                        problems.append("输入法插件补丁未能自动修复")
                        print(out.strip()[-400:])
                else:
                    if use_bundled:
                        problems.append("输入法插件补丁未应用")
                    else:
                        # 当前用的是系统同版本插件，它本就不需要这套适配补丁，
                        # 不能据此报错，否则会在 x86_64 / 新版 Qt 上误报。
                        print("  说明：当前使用系统 Qt6 输入法插件（版本自洽），无需本适配补丁。")
        else:
            print("  提示：缺少 tools/patch_qt6_ime_plugin.py，无法自检插件补丁。")

    # ---------- 2) libQt6Gui ----------
    # 该补丁只对「按 Qt 6.4 编译的插件 + Qt 6.7 宿主」这一组合有意义。
    # 其它 Qt 版本（例如系统自带、版本自洽的 PyQt6）不应套用，也不应据此报错。
    tool_gui = _tool(root, "patch_qt6_inputcontext.py")
    need_gui_patch = bool(ver and ver.startswith(TARGET_QT_PREFIX))
    if os.path.isfile(gui_lib) and tool_gui and need_gui_patch:
        rc, out = _run([sys.executable, tool_gui, "--check", gui_lib])
        patched = ("已修补" in out)
        print("libQt6Gui：%s" % ("已打补丁" if patched else "未打补丁"))
        if not patched:
            if a.repair:
                rc2, out2 = _run([sys.executable, tool_gui, gui_lib])
                if rc2 == 0:
                    rc3, out3 = _run([sys.executable, tool_gui, "--check", gui_lib])
                    print("  → 已应用补丁，复核：%s" % ("OK" if "已修补" in out3 else "仍异常"))
                else:
                    problems.append("libQt6Gui 补丁应用失败")
                    print(out2.strip()[-400:])
            else:
                problems.append("libQt6Gui 未打补丁（会导致界面无法输入中文并崩溃）")
    elif os.path.isfile(gui_lib) and not need_gui_patch:
        print("libQt6Gui：Qt %s 与插件版本自洽，无需本适配补丁" % (ver or "未知"))
    elif not os.path.isfile(gui_lib):
        print("libQt6Gui：未找到 %s" % gui_lib)

    print("=" * 62)
    if problems:
        print("存在问题：")
        for p in problems:
            print("  - %s" % p)
        print("=" * 62)
        return 1
    if not setup:
        print("结论：需系统提供 Qt6 输入法插件（见上方安装命令）。")
        print("=" * 62)
        return 0
    print("结论：中文输入相关补丁就绪。")
    print("=" * 62)
    return 0


if __name__ == "__main__":
    sys.exit(main())
