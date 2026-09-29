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

# 兜底：部分历史包（如 Windows 便携版基线）不含 app/ui/appicon.py。
# 顶层导入失败会让程序根本起不来，所以这里退化为内置常量，保证入口永远可用。
try:
    from app.ui import appicon
except Exception:  # noqa: BLE001
    appicon = None
from app.ui.main_window import main

_APP_ID = getattr(appicon, "APP_ID", "wenhui-assembler")


def _adopt_desktop_identity():
    """把 argv[0] 换成应用 ID，使窗口身份与桌面项文件名严格一致。

    背景（实测定位）：Qt 在 xcb 下把 WM_CLASS 写成
    ``"<argv[0] 的 basename>", "<applicationName>"``。直接用 ``python3 app/main.py``
    启动时 argv[0] 是 ``main.py``，于是 instance 段变成 ``main.py``。

    而 UKUI 的 panel-daemon（ukui-panel 3.25）**根本不读 StartupWMClass**（其二进制里
    没有该字符串）。它按「窗口 app-id / 进程 cmdline ↔ 桌面项文件名」匹配，匹配失败
    就回退到进程名。系统里没有 ``main.py.desktop``，于是回退读到 python3，命中系统自带
    的 ``/usr/share/applications/python3.8.desktop`` → 任务栏显示 **Python 图标**。

    把 argv[0] 置为应用 ID 后（实测 Xvfb + 本包自带 PyQt6 6.7.3）：

        WM_CLASS = ("wenhui-assembler", "wenhui-assembler")
        _GTK_APPLICATION_ID      = "wenhui-assembler"
        _KDE_NET_WM_DESKTOP_FILE = "wenhui-assembler"

    三个来源（WM_CLASS 两段、GTK AppId、KDE 桌面项）全部指向
    ``~/.local/share/applications/wenhui-assembler.desktop``，无论桌面环境读哪一个
    都能取到自定义图标。启动脚本另用 ``exec -a`` 让 ``/proc/<pid>/cmdline`` 的
    argv[0] 同为应用 ID，覆盖面板按 cmdline 回退匹配的路径。

    必须在本进程创建 QApplication 之前执行（WM_CLASS 只在窗口创建时写一次）。
    """
    try:
        sys.argv[0] = _APP_ID
    except Exception:  # noqa: BLE001
        pass


if __name__ == "__main__":
    _adopt_desktop_identity()
    main()
