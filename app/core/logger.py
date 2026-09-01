# -*- coding: utf-8 -*-
"""日志模块：把程序运行日志写入文件，便于排查问题。

日志文件默认位于程序根目录下的 logs/assembler.log，单文件 2MB 自动轮转，
保留最近 5 份。若程序目录不可写（如只读盘），自动回退到系统临时目录。
"""
import logging
import os
import sys
import tempfile
from logging.handlers import RotatingFileHandler

_configured = False
LOG_FILE = None


def _project_root():
    # logger.py 位于 app/core/，向上三级即项目根目录
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _resolve_log_dir():
    """优先用程序目录 logs/，不可写则回退系统临时目录，均失败返回 None。"""
    candidates = [os.path.join(_project_root(), "logs")]
    try:
        candidates.append(os.path.join(tempfile.gettempdir(), "doc_assembler_logs"))
    except Exception:
        pass
    for d in candidates:
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, ".probe")
            with open(probe, "w", encoding="utf-8") as f:
                f.write("ok")
            os.remove(probe)
            return d
        except Exception:
            continue
    return None


def setup_logging(level=logging.INFO):
    """初始化日志（幂等）。返回日志文件绝对路径，可能为 None。"""
    global _configured, LOG_FILE
    if _configured:
        return LOG_FILE
    _configured = True

    log_dir = _resolve_log_dir()
    root = logging.getLogger()
    root.setLevel(level)

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    if log_dir:
        LOG_FILE = os.path.join(log_dir, "assembler.log")
        fh = RotatingFileHandler(LOG_FILE, maxBytes=2 * 1024 * 1024,
                                 backupCount=5, encoding="utf-8")
        fh.setFormatter(fmt)
        root.addHandler(fh)

    # 控制台输出：以 pythonw 运行时无控制台，捕获异常静默跳过
    try:
        sh = logging.StreamHandler(sys.stderr)
        sh.setFormatter(fmt)
        root.addHandler(sh)
    except Exception:
        pass

    # 未捕获异常也写入日志
    def _excepthook(etype, value, tb):
        root.critical("未捕获异常", exc_info=(etype, value, tb))
    sys.excepthook = _excepthook

    root.info("日志系统初始化完成，日志文件：%s", LOG_FILE or "（无）")
    return LOG_FILE


def get_logger(name):
    return logging.getLogger(name)
