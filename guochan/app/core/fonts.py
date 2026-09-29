# -*- coding: utf-8 -*-
"""
fonts.py —— 字体检测
必需字体：方正小标宋简体 / 仿宋_GB2312 / 楷体_GB2312 / 黑体 / 宋体(SimSun)
依赖系统已安装；若缺失则提示用户安装（随附字体包位于 app/fonts_bundled/）。
"""
import os
import platform

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUNDLED_FONT_DIR = os.path.join(APP_DIR, "fonts_bundled")

REQUIRED = {
    "方正小标宋简体": ["fzxiao", "小标宋", "fangzhengxiaobiaosong"],
    "仿宋_GB2312": ["仿宋_gb2312", "fangsong", "fangsong_gb2312"],
    "楷体_GB2312": ["楷体_gb2312", "kaiti", "kaiti_gb2312"],
    "黑体": ["simhei", "黑体", "heiti"],
    "宋体": ["simsun", "宋体", "songti"],
}


def _system_font_names():
    """尽力收集系统已安装字体族名/文件名（小写）。"""
    names = set()
    # Linux: fc-list
    if platform.system() != "Windows":
        import shutil
        if shutil.which("fc-list"):
            try:
                import subprocess
                out = subprocess.run(["fc-list"], capture_output=True, text=True, timeout=30).stdout
                for line in out.splitlines():
                    # fc-list 格式: 路径: 族名: 样式
                    if ":" in line:
                        fam = line.split(":", 1)[1].split(":")[0].strip().lower()
                        names.add(fam)
                    names.add(os.path.basename(line.split(":")[0]).lower())
            except Exception:
                pass
    else:
        # Windows: 扫描 C:\Windows\Fonts
        fd = r"C:\Windows\Fonts"
        if os.path.isdir(fd):
            for f in os.listdir(fd):
                names.add(f.lower())
    return names


def check_fonts():
    """
    返回 (missing_list, bundled_present)
      missing_list: 可能缺失的必需字体名（best-effort）
      bundled_present: 随附字体包目录是否存在
    """
    sys_names = _system_font_names()
    missing = []
    for fam, keys in REQUIRED.items():
        ok = any(any(k in n for n in sys_names) for k in keys)
        if not ok:
            missing.append(fam)
    bundled = os.path.isdir(BUNDLED_FONT_DIR) and any(
        f.lower().endswith((".ttf", ".ttc", ".otf")) for f in os.listdir(BUNDLED_FONT_DIR)
    )
    return missing, bundled


def bundled_font_dir():
    return BUNDLED_FONT_DIR if os.path.isdir(BUNDLED_FONT_DIR) else None
