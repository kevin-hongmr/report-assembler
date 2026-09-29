# -*- coding: utf-8 -*-
"""程序入口。"""
import os
import sys

# main.py 位于 app/ 内，要让 `app.ui...` 可被导入，需把项目根目录加入 sys.path
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

# 先初始化日志，保证后续所有运行日志写入文件
from app.core import logger as logger_mod
logger_mod.setup_logging()

from app.ui.main_window import main

if __name__ == "__main__":
    main()
