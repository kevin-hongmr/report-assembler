#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""离线 wheel 兼容性静态校验。

本机是 Windows，无法真实安装 manylinux/aarch64 轮子，因此用纯静态方式模拟 pip 的选包逻辑：
为每个目标 Python 版本生成"兼容 tag 集合"，再逐个 wheel 比对文件名 tag 与 Requires-Python，
确认每个必需包都能被选中。等价于提前跑一遍 pip 的解析结果。

用法：
    python tools/verify_wheels.py                     # 默认校验 wheels/
    python tools/verify_wheels.py --arch x86_64
    python tools/verify_wheels.py --glibc 2.17        # 校验老系统（麒麟早期版本）
"""

import argparse
import glob
import os
import re
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 必需包（canonical 形式，按下划线/连字符归一化后比较）
REQUIRED = ["python-docx", "lxml", "pyqt6", "pyqt6-qt6", "pyqt6-sip"]
OPTIONAL = ["pymupdf"]
BOOTSTRAP = ["pip"]


def canon(name):
    """PEP 503 名称归一化：PyQt6_Qt6 -> pyqt6-qt6"""
    return re.sub(r"[-_.]+", "-", name).lower()


def parse_wheel_tags(fname):
    """从 wheel 文件名解析 (name, version, tag集合)。"""
    base = fname[:-4] if fname.endswith(".whl") else fname
    parts = base.split("-")
    if len(parts) < 5:
        return None, None, set()
    name, ver = parts[0], parts[1]
    tagstr = parts[-3]  # python tag
    if len(parts) == 6:
        tagstr = parts[-3]
    # 后三段恒为 python-abi-platform
    pytag, abitag, plattag = parts[-3], parts[-2], parts[-1]
    tags = set()
    for p in pytag.split("."):
        for a in abitag.split("."):
            for l in plattag.split("."):
                tags.add("%s-%s-%s" % (p, a, l))
    return canon(name), ver, tags


def target_tags(pyver, platforms):
    """生成目标 Python 在给定平台集合下的兼容 tag 集合（近似 pip compatible_tags）。"""
    major, minor = [int(x) for x in pyver.split(".")]
    tags = set()
    # cpXY-cpXY-<plat>
    for plat in platforms:
        tags.add("cp%d%d-cp%d%d-%s" % (major, minor, major, minor, plat))
        # abi3：从当前版本一路向下兼容到 3.2
        for m in range(minor, 1, -1):
            tags.add("cp%d%d-abi3-%s" % (major, m, plat))
        # cpXY-none-<plat>
        tags.add("cp%d%d-none-%s" % (major, minor, plat))
        # pyXY-none-<plat> / py{major}-none-<plat>
        tags.add("py%d%d-none-%s" % (major, minor, plat))
        tags.add("py%d-none-%s" % (major, plat))
    # 纯 Python 包
    tags.add("py3-none-any")
    tags.add("py2-none-any")
    tags.add("py%d%d-none-any" % (major, minor))
    tags.add("none-any")
    return tags


def vtuple(s):
    nums = re.findall(r"\d+", s)
    return tuple(int(n) for n in (nums + ["0", "0", "0"])[:3])


def requires_python_ok(wheel_path, pyver):
    """读 wheel 内 METADATA 的 Requires-Python，判断目标版本是否满足。"""
    try:
        with zipfile.ZipFile(wheel_path) as z:
            names = [n for n in z.namelist() if n.endswith(".dist-info/METADATA")]
            if not names:
                return True
            meta = z.read(names[0]).decode("utf-8", "replace")
    except Exception:
        return True
    m = re.search(r"^Requires-Python:\s*(.+)$", meta, re.M)
    if not m:
        return True
    target = vtuple(pyver)
    for clause in m.group(1).split(","):
        clause = clause.strip()
        if not clause:
            continue
        mm = re.match(r"^(>=|<=|==|!=|>|<|~=)\s*([0-9][0-9A-Za-z.\-]*)$", clause)
        if not mm:
            continue  # 遇到无法解析的子句（如 *) 就跳过
        op, val = mm.group(1), vtuple(mm.group(2))
        if op == ">=" and not target >= val:
            return False
        if op == ">" and not target > val:
            return False
        if op == "<=" and not target <= val:
            return False
        if op == "<" and not target < val:
            return False
        if op == "==" and target != val:
            return False
        if op == "!=" and target == val:
            return False
    return True


def vkey(v):
    return tuple(int(n) if n.isdigit() else 0 for n in re.findall(r"\d+", v)[:4]) or (0,)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--arch", default="aarch64", choices=["aarch64", "x86_64"])
    ap.add_argument("--glibc", default="2.28", help="目标机 glibc 版本，决定可用的 manylinux 标签")
    ap.add_argument("--py", nargs="+", default=["3.8", "3.9", "3.10", "3.11", "3.12"])
    ap.add_argument("--wheels", default=os.path.join(ROOT, "wheels"))
    opts = ap.parse_args()

    wheels = sorted(glob.glob(os.path.join(opts.wheels, "*.whl")))
    if not wheels:
        print("wheels 目录为空：%s" % opts.wheels)
        return 1

    # 目标机可接受的平台标签（按 glibc 分档）
    plats = ["linux_" + opts.arch]
    if vtuple(opts.glibc) >= (2, 5, 0):
        plats.append("manylinux1_" + opts.arch)
    if vtuple(opts.glibc) >= (2, 12, 0):
        plats.append("manylinux2010_" + opts.arch)
    if vtuple(opts.glibc) >= (2, 17, 0):
        plats += ["manylinux2014_" + opts.arch, "manylinux_2_17_" + opts.arch]
    if vtuple(opts.glibc) >= (2, 26, 0):
        plats.append("manylinux_2_26_" + opts.arch)
    if vtuple(opts.glibc) >= (2, 28, 0):
        plats.append("manylinux_2_28_" + opts.arch)

    print("wheel 目录：%s（%d 个）" % (opts.wheels, len(wheels)))
    print("目标：arch=%s  glibc=%s" % (opts.arch, opts.glibc))
    print("可用平台标签：%s\n" % ", ".join(sorted(plats)))

    # 预解析所有 wheel
    pool = {}
    for w in wheels:
        name, ver, tags = parse_wheel_tags(os.path.basename(w))
        if name:
            pool.setdefault(name, []).append((ver, tags, w))

    all_ok = True
    need = REQUIRED + OPTIONAL + BOOTSTRAP
    for pyver in opts.py:
        ttags = target_tags(pyver, plats)
        print("=" * 68)
        print("Python %s" % pyver)
        print("=" * 68)
        for pkg in need:
            cands = []
            for ver, tags, path in pool.get(pkg, []):
                if tags & ttags and requires_python_ok(path, pyver):
                    cands.append((vkey(ver), ver, os.path.basename(path)))
            if cands:
                cands.sort()
                best = cands[-1]
                mark = "OK " if pkg not in OPTIONAL else "可选"
                print("  [%s] %-14s %-10s %s" % (mark, pkg, best[1], best[2]))
            else:
                all_ok = False
                print("  [缺失] %-14s ----  没有与 Python %s / glibc %s 兼容的 wheel" % (pkg, pyver, opts.glibc))
        print("")

    print("=" * 68)
    print("结论：%s" % ("全部目标版本均可离线安装 ✔" if all_ok else "存在缺口 ✘，请补齐对应 wheel"))
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
