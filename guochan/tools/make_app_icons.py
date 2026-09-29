#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从原始大图生成程序所需的多尺寸图标，写入 assets/icons/。

## 为什么必须生成多尺寸，不能直接用原始大图

实测（Qt 6.7.3 / xcb，xprop 读取窗口属性）：

| QIcon 构成                       | _NET_WM_ICON |
|----------------------------------|--------------|
| 仅 1641x1641 单张（原始大图）     | **空**       |
| 仅 256x256 单张                   | **空**       |
| 16+48 两张                        | 有数据       |
| 16+32+48+64+128+256 七张          | 有数据       |
| 七张 + 512                        | 有数据       |
| 1641 大图 + 若干小图              | **空**       |

结论：Qt6 在 xcb 平台写入 `_NET_WM_ICON` 时，
1. 需要 QIcon 中存在**至少 2 个**位图；
2. 单个位图尺寸过大会导致整体序列化失败（实测 1641 会使结果为空；<=512 正常）。

因此程序图标必须由「多张、每张不超过 512px」的 PNG 组成。原始 `图标.png`
为 1641x1641，直接 `QIcon(图标.png)` 会使任务栏拿不到窗口图标，
桌面环境只能回退去匹配进程（python3）的图标——表现为「任务栏显示 Python 图标」。

用法：
    python3 tools/make_app_icons.py                 # 源图 图标.png，输出 assets/icons/
    python3 tools/make_app_icons.py --src <图> --out <目录>
"""

import argparse
import os
import shutil
import subprocess
import sys

# 覆盖桌面环境与任务栏常见的请求尺寸；上限 256 已足够（512 亦可，但无必要）
DEFAULT_SIZES = (16, 24, 32, 48, 64, 128, 256)


def _make_with_pil(src, out, sizes):
    try:
        from PIL import Image  # type: ignore
    except ImportError:
        return False
    im = Image.open(src).convert("RGBA")
    if im.width != im.height:
        # 非正方形：居中裁剪为正方形，避免缩放变形
        side = min(im.width, im.height)
        left = (im.width - side) // 2
        top = (im.height - side) // 2
        im = im.crop((left, top, left + side, top + side))
    for s in sizes:
        im.resize((s, s), Image.LANCZOS).save(os.path.join(out, "%d.png" % s), optimize=True)
    return True


def _make_with_imagemagick(src, out, sizes):
    conv = shutil.which("convert") or shutil.which("magick")
    if not conv:
        return False
    ok = True
    for s in sizes:
        dst = os.path.join(out, "%d.png" % s)
        r = subprocess.run([conv, src, "-resize", "%dx%d" % (s, s), "-gravity", "center",
                            "-extent", "%dx%d" % (s, s), dst],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        ok = ok and r.returncode == 0
    return ok


def _make_with_qt(src, out, sizes):
    """最后兜底：用 PyQt6 自身缩放（目标机上通常必有 PyQt6）。"""
    try:
        from PyQt6.QtCore import Qt
        from PyQt6.QtGui import QImage
    except ImportError:
        return False
    img = QImage(src)
    if img.isNull():
        return False
    for s in sizes:
        img.scaled(s, s, Qt.AspectRatioMode.IgnoreAspectRatio,
                   Qt.TransformationMode.SmoothTransformation).save(os.path.join(out, "%d.png" % s), "PNG")
    return True


def main():
    ap = argparse.ArgumentParser()
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--src", default=os.path.join(root, "图标.png"))
    ap.add_argument("--out", default=os.path.join(root, "assets", "icons"))
    ap.add_argument("--sizes", default=",".join(str(s) for s in DEFAULT_SIZES))
    a = ap.parse_args()

    sizes = tuple(int(x) for x in a.sizes.split(",") if x.strip())
    if not os.path.isfile(a.src):
        print("错误：找不到源图 %s" % a.src)
        return 1
    os.makedirs(a.out, exist_ok=True)

    for name, fn in (("Pillow", _make_with_pil), ("ImageMagick", _make_with_imagemagick),
                     ("PyQt6", _make_with_qt)):
        try:
            if fn(a.src, a.out, sizes):
                print("使用 %s 生成多尺寸图标：%s" % (name, ", ".join(str(s) for s in sizes)))
                break
        except Exception as e:  # noqa: BLE001
            print("  %s 生成失败：%s" % (name, e))
    else:
        print("错误：Pillow / ImageMagick / PyQt6 均不可用，无法生成图标。")
        print("      请安装其一：pip install Pillow  或  sudo apt install imagemagick")
        return 1

    missing = [s for s in sizes if not os.path.isfile(os.path.join(a.out, "%d.png" % s))]
    if missing:
        print("警告：以下尺寸未生成：%s" % missing)
    print("输出目录：%s" % a.out)
    for s in sizes:
        p = os.path.join(a.out, "%d.png" % s)
        if os.path.isfile(p):
            print("  %-8s %8d 字节" % ("%dx%d" % (s, s), os.path.getsize(p)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
