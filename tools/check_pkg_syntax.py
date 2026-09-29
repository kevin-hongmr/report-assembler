#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把各交付包内 app/ 下的 Python 源码在指定解释器上做语法编译校验（不落盘、不执行）。

用途：交付包要同时在「Windows 内置 Python 3.13」与「麒麟系统 Python 3.8」上运行，
而本机只能跑其中一个平台。语法编译是少数**能跨平台验证**的检查项：
既能把版本不兼容的语法挡下来（如只在 3.9+ 的写法），也能抓住"文件搬运搬坏了"这类事故。

用法：
    python3 check_pkg_syntax.py <zip> [<zip> ...]
"""
import sys
import zipfile

SKIP_PREFIX = ("app/runtime/", "app/libreoffice/", "app/.venv/")


def check(zpath):
    zf = zipfile.ZipFile(zpath)
    files = []
    for n in zf.namelist():
        if not n.endswith(".py"):
            continue
        if "/app/" not in n:
            continue
        rel = n.split("/app/", 1)[1]
        if any(rel.startswith(p[4:]) for p in SKIP_PREFIX):
            continue
        if rel.startswith("app/"):
            continue
        files.append((n, rel))
    ok, bad = 0, []
    for n, rel in sorted(files):
        try:
            compile(zf.read(n), rel, "exec")
            ok += 1
        except SyntaxError as e:
            bad.append("%s:%s %s" % (rel, e.lineno, e.msg))
        except Exception as e:  # noqa: BLE001
            bad.append("%s %s" % (rel, e))
    return ok, bad


def main():
    ver = "%d.%d.%d" % sys.version_info[:3]
    print("=" * 72)
    print("解释器 Python %s" % ver)
    print("=" * 72)
    rc = 0
    for z in sys.argv[1:]:
        ok, bad = check(z)
        print("%-46s 通过 %3d 个 .py" % (z, ok))
        for b in bad:
            print("     ★ 语法错误：%s" % b)
            rc = 1
    print("结论：", "全部语法合法 ✔" if rc == 0 else "★ 存在语法错误")
    return rc


if __name__ == "__main__":
    sys.exit(main())
