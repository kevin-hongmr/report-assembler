#!/usr/bin/env bash
# ============================================
#  批量把 .doc/.wps 转成 .docx（Windows 用 WPS；国产 Linux 用 LibreOffice）
#  用法：把 .doc/.wps 文件或文件夹拖到本脚本上（或命令行传参）即批量转换。
#  转换后的 .docx 输出在原位置，可直接添加进汇编程序。
# ============================================
cd "$(dirname "$0")/.." || exit 1
PY="app/.venv/bin/python"
[ -x "$PY" ] || PY="python3"

if [ $# -eq 0 ]; then
  echo "请把 .doc/.wps 文件或文件夹拖到本脚本上；"
  echo "或命令行运行：$0 文件或文件夹 [--outdir 目录]"
  echo
  "$PY" tools/convert_to_docx.py --help
  exit 0
fi

"$PY" tools/convert_to_docx.py "$@"
