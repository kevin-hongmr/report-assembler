# -*- coding: utf-8 -*-
"""
convert_to_docx.py —— 批量把 .doc / .wps 转成 .docx（独立转换脚本）

用法：
  python convert_to_docx.py 文件1 [文件2 ...] [--outdir 输出目录]
  python convert_to_docx.py 文件夹            # 递归转换该目录下所有 .doc/.wps
  直接把 .doc/.wps 文件或文件夹拖到「转docx.bat」/「转docx.sh」上即可。

转换引擎（由 app/core/converter 自动探测，无需手动选择）：
  - Windows：WPS COM 自动化（个人版 / 专业版均可，无需命令行支持）
  - 麒麟 / 统信等国产 Linux：LibreOffice 无头转换
    （WPS for Linux 不支持命令行转换，故国产系统走 LibreOffice）

转换完成后，把生成的 .docx 添加进「文档汇编程序」即可汇编（.docx 汇编不依赖任何引擎）。
"""
import os
import sys
import glob
import shutil
import argparse
import tempfile

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from app.core import converter


def collect(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for ext in (".doc", ".docx", ".wps"):
                files += sorted(glob.glob(os.path.join(p, "*" + ext)))
                files += sorted(glob.glob(os.path.join(p, "**", "*" + ext), recursive=True))
        elif os.path.isfile(p):
            files.append(p)
    out = []
    for f in files:
        if os.path.basename(f).startswith("~$"):
            continue  # 跳过 WPS/Word 打开文档时生成的临时锁文件
        if f.lower().endswith((".doc", ".docx", ".wps")):
            out.append(os.path.abspath(f))
    return sorted(set(out))


def main():
    ap = argparse.ArgumentParser(description="把 .doc/.wps 批量转换为 .docx")
    ap.add_argument("paths", nargs="*", help="文件或目录（支持拖拽到 bat/sh）")
    ap.add_argument("--outdir", default=None, help="输出目录（默认与源文件同目录）")
    args = ap.parse_args()
    if not args.paths:
        print(__doc__)
        sys.exit(1)
    srcs = collect(args.paths)
    if not srcs:
        print("未找到 .doc/.docx/.wps 文件。")
        sys.exit(1)
    print(f"待处理 {len(srcs)} 个文件：")
    work = tempfile.mkdtemp(prefix="conv_batch_")
    ok = fail = skip = 0
    for s in srcs:
        # 已是干净 .docx 的（含普通 docx）无需转换，跳过
        if s.lower().endswith(".docx") and not converter._docx_needs_normalize(s):
            print(f"  跳过（已是干净 docx） {os.path.basename(s)}")
            skip += 1
            continue
        try:
            dx = converter.convert_to_docx(s, work)
            dst_dir = os.path.abspath(args.outdir) if args.outdir else os.path.dirname(s)
            os.makedirs(dst_dir, exist_ok=True)
            dst = os.path.join(dst_dir, os.path.splitext(os.path.basename(s))[0] + ".docx")
            shutil.copy2(dx, dst)
            print(f"  OK  {os.path.basename(s)}  ->  {os.path.basename(dst)}")
            ok += 1
        except Exception as e:
            print(f"  FAIL {os.path.basename(s)}: {e}")
            fail += 1
    print(f"\n完成：成功 {ok}，失败 {fail}，跳过 {skip}（已为干净 docx）。"
          f"转换后的 .docx 可直接添加进汇编程序。")
    sys.exit(0 if fail == 0 else 2)


if __name__ == "__main__":
    main()
