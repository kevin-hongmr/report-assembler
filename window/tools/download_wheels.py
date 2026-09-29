#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""为国产 Linux（银河麒麟 / 统信 UOS）下载离线依赖 wheel 到 wheels/ 目录。

设计要点
--------
1. PyQt6 / PyQt6-Qt6 / PyMuPDF / python-docx / pip 等轮子是 abi3 或 py3-none 标签，
   一份即可被 Python 3.8 ~ 3.12 共用，只随 CPU 架构变化，因此只下载一次。
2. lxml / PyQt6-sip 是 cp3X-cp3X 强绑定轮子，必须按目标 Python 版本逐一下载。
3. 平台标签同时给出 manylinux_2_28 / manylinux_2_17 / manylinux1，
   这样 PyQt6-Qt6（只有 2_28）与 PyQt6-sip（只有 2_17）都能命中。

用法
----
    python tools/download_wheels.py                          # 默认 aarch64 + 3.8~3.12
    python tools/download_wheels.py --arch x86_64
    python tools/download_wheels.py --arch aarch64 --py 3.8 3.9
"""

import argparse
import os
import shutil
import subprocess
import sys
import glob

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WHEELS = os.path.join(ROOT, "wheels")

# 平台标签：从新到旧，pip 会按兼容性挑选
PLATFORM_TAGS = {
    "aarch64": ["manylinux_2_28_aarch64", "manylinux_2_17_aarch64", "manylinux1_aarch64"],
    "x86_64": ["manylinux_2_28_x86_64", "manylinux_2_17_x86_64", "manylinux1_x86_64"],
}

# 与 Python 版本无关的共享包（abi3 / py3-none）
SHARED = ["PyQt6", "PyMuPDF", "python-docx", "pip", "setuptools"]
# 与 Python 版本强绑定的包（cp3X-cp3X）
PER_VERSION = ["lxml", "PyQt6-sip"]

DEFAULT_PY = ["3.8", "3.9", "3.10", "3.11", "3.12"]


def pip_download(args, target_dir):
    """调用 pip download，返回 (ok, output)。"""
    cmd = [sys.executable, "-m", "pip", "download",
           "--only-binary=:all:",
           "--no-cache-dir",
           "-d", target_dir]
    cmd += args
    print("  $ " + " ".join(cmd))
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = (proc.stdout or "") + (proc.stderr or "")
    return proc.returncode == 0, out


def build_platform_args(pyver, arch):
    args = ["--python-version", pyver,
            "--implementation", "cp",
            "--abi", "cp" + pyver.replace(".", "")]
    for plat in PLATFORM_TAGS[arch]:
        args += ["--platform", plat]
    return args


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--arch", default="aarch64", choices=list(PLATFORM_TAGS))
    ap.add_argument("--py", nargs="+", default=DEFAULT_PY, help="目标 Python 版本列表")
    ap.add_argument("--clean", action="store_true", help="下载前清空输出目录")
    ap.add_argument("--out", default=None,
                    help="输出目录，默认 wheels/。建议按架构分开存放，如 wheels_aarch64/、"
                         "wheels_x86_64/，避免两个架构的 wheel 互相覆盖")
    opts = ap.parse_args()

    wheels_dir = os.path.abspath(opts.out) if opts.out else WHEELS
    if opts.clean and os.path.isdir(wheels_dir):
        shutil.rmtree(wheels_dir)
    os.makedirs(wheels_dir, exist_ok=True)

    base_ver = opts.py[0]  # 用最低版本下共享包，兼容性最广

    print("=" * 70)
    print("架构：%s    目标 Python：%s" % (opts.arch, " / ".join(opts.py)))
    print("输出：%s" % wheels_dir)
    print("=" * 70)

    # ---- 1. 共享包（只下一次，用最低 Python 版本） ----
    print("\n[1/3] 下载共享包（abi3 / py3-none，全版本共用）")
    ok, out = pip_download(build_platform_args(base_ver, opts.arch) + SHARED, wheels_dir)
    print(out.strip()[-2500:] if not ok else "  完成")
    if not ok:
        print("!! 共享包下载失败，请检查网络后重试")
        return 1

    # ---- 2. 按版本强绑定包 ----
    print("\n[2/3] 下载按版本强绑定包（lxml / PyQt6-sip）")
    failed = []
    for pyver in opts.py:
        print("  -- Python %s --" % pyver)
        ok, out = pip_download(build_platform_args(pyver, opts.arch) + PER_VERSION, wheels_dir)
        if ok:
            n = len([l for l in out.splitlines() if l.startswith("Saved ") or "Saved" in l])
            print("     完成（本次新增 %d 个）" % n)
        else:
            print("     失败：%s" % out.strip()[-400:])
            failed.append(pyver)
    if failed:
        print("!! 以下 Python 版本下载失败：%s" % ", ".join(failed))

    # ---- 3. 清理与汇总 ----
    print("\n[3/3] 汇总")
    files = sorted(glob.glob(os.path.join(wheels_dir, "*.whl")))
    total = sum(os.path.getsize(f) for f in files)
    for f in files:
        print("     %8.1f MB  %s" % (os.path.getsize(f) / 1048576.0, os.path.basename(f)))
    print("\n共 %d 个 wheel，合计 %.1f MB" % (len(files), total / 1048576.0))

    # 校验每个目标版本都有 lxml 与 sip
    print("\n版本覆盖校验：")
    for pyver in opts.py:
        tag = "cp" + pyver.replace(".", "")
        has_lxml = glob.glob(os.path.join(wheels_dir, "lxml-*-%s-%s-*.whl" % (tag, tag)))
        has_sip = glob.glob(os.path.join(wheels_dir, "PyQt6_sip-*-%s-%s-*.whl" % (tag, tag)))
        print("  Python %-5s lxml:%-3s PyQt6-sip:%-3s" % (pyver, "OK" if has_lxml else "缺",
                                                          "OK" if has_sip else "缺"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
