# -*- coding: utf-8 -*-
"""日志模块：把程序运行日志写入文件，便于排查问题。

日志文件默认位于程序根目录下的 logs/assembler.log，单文件 2MB 自动轮转，
保留最近 5 份。若程序目录不可写（如只读盘），自动回退到系统临时目录。

除常规日志外，本模块还负责两件"排障兜底"的事：

1. ``sys.excepthook`` —— Python 层未捕获异常（主线程 / 槽函数）写入日志；
2. ``faulthandler`` —— **C 层崩溃取证**。SIGSEGV / SIGABRT 之类的硬崩溃不会经过
   Python 异常机制，日志里只会突兀地断在最后一行（"闪退"往往就是这种）。启用
   faulthandler 后，崩溃瞬间会把各线程的 Python 调用栈写进同一个日志文件，
   下次再崩就能直接看出卡在哪个调用上。
"""
import faulthandler
import logging
import os
import sys
import tempfile
import threading
from logging.handlers import RotatingFileHandler

_configured = False
LOG_FILE = None
# 必须保持强引用，否则崩溃时文件已被 GC 关闭
_CRASH_STREAM = None


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

    # 子线程里的未捕获异常同样写入日志（默认只打到 stderr，而 GUI 启动时
    # stderr 常被丢弃）
    try:
        def _thread_excepthook(args):
            root.critical("子线程未捕获异常：%s",
                          args.exc_value, exc_info=(args.exc_type, args.exc_value,
                                                    args.exc_traceback))
        threading.excepthook = _thread_excepthook
    except Exception:  # noqa: BLE001
        pass

    # C 层崩溃取证（SIGSEGV / SIGABRT …）：把崩溃时的 Python 调用栈写进同一日志。
    # SIGSEGV 走不到 excepthook，没有这段的话日志会突兀断掉、无从定位。
    global _CRASH_STREAM
    try:
        if LOG_FILE:
            _CRASH_STREAM = open(LOG_FILE, "a", encoding="utf-8", buffering=1)
            faulthandler.enable(file=_CRASH_STREAM, all_threads=True)
    except Exception:  # noqa: BLE001
        pass

    root.info("日志系统初始化完成，日志文件：%s", LOG_FILE or "（无）")
    return LOG_FILE


def get_logger(name):
    return logging.getLogger(name)
