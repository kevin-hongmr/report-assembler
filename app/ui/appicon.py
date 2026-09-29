# -*- coding: utf-8 -*-
"""程序身份与桌面集成：让任务栏显示自定义图标而非 Python 图标。

## 问题成因（已实测定位）

在 Linux 桌面上，任务栏/停靠栏解析「某个窗口该用哪个图标」的顺序大致是：

1. 先用窗口的 `WM_CLASS` 的 class 段去匹配已安装的 `.desktop` 文件（`StartupWMClass=`）；
2. 匹配不到时，退而用进程名 / AppId 去猜 —— 本程序由 `python3 app/main.py` 启动，
   于是被认成 `python3`，任务栏就显示了 **Python 的图标**。

实测原始状态（Xvfb + xprop）：

    WM_CLASS(STRING) = "main.py", "main.py"
    _NET_WM_ICON(CARDINAL) =            ← 空
    _GTK_APPLICATION_ID = "python3"
    _KDE_NET_WM_DESKTOP_FILE = "python3"

即两个问题同时存在：

- **没有可匹配的 .desktop**：`WM_CLASS` 是 `main.py`，系统里没有对应的桌面项，
  图标解析只能回退到进程 `python3`；
- **窗口根本没有携带图标数据**：`_NET_WM_ICON` 为空。原因是原代码用
  `QIcon("图标.png")`（单张 1641x1641），而 Qt6 在 xcb 下写入 `_NET_WM_ICON`
  时要求 QIcon 含**至少 2 个位图**且**单个尺寸不能过大**（实测 ≥1641 会导致整体为空，
  ≤512 正常）。详见 `tools/make_app_icons.py` 中的实测表。

## 本模块做的事

- `apply_app_identity(app)`：设置 applicationName / applicationDisplayName /
  desktopFileName，并用**多尺寸** PNG 组装 QIcon 后设置窗口图标；
- `install_desktop_entry()`：把 `.desktop` 与多尺寸图标安装到用户目录
  （`~/.local/share/applications` 与 `~/.local/share/icons/hicolor/*/apps`），
  其中 `StartupWMClass` 与 `applicationName` 保持一致，供桌面环境据此匹配图标。

以上全部为「尽力而为」：任何一步失败都只记日志，绝不影响程序启动。
"""

import os
import shutil
import subprocess
import sys

# 程序在桌面环境中的唯一标识。须与 .desktop 文件名、StartupWMClass 一致。
APP_ID = "wenhui-assembler"
APP_NAME = "文档汇编程序"
APP_DISPLAY_NAME = "文档汇编程序 · 政务版"
APP_COMMENT = "将多份 Word 文档汇编为符合公文规范的文档（自动生成封面与目录）"
APP_ORG = "ChaozhouGov"

# 必须与 tools/make_app_icons.py 的 DEFAULT_SIZES 一致；上限不超过 512
ICON_SIZES = (16, 24, 32, 48, 64, 128, 256)


def _log():
    try:
        from app.core.logger import get_logger
        return get_logger(__name__)
    except Exception:  # noqa: BLE001
        import logging
        return logging.getLogger(__name__)


def project_root():
    """本文件位于 app/ui/，向上三级即项目根目录。"""
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def icons_dir():
    """定位多尺寸图标目录；兼容项目根 assets/icons 与 app/assets/icons 两种布局。"""
    root = project_root()
    for rel in (os.path.join("assets", "icons"), os.path.join("app", "assets", "icons")):
        d = os.path.join(root, rel)
        if os.path.isdir(d) and any(f.endswith(".png") for f in os.listdir(d)):
            return d
    return None


def legacy_icon_path():
    """回退用：原始大图。仅在多尺寸图标缺失时使用（届时窗口图标可能为空）。"""
    root = project_root()
    for rel in ("图标.png", os.path.join("app", "图标.png")):
        p = os.path.join(root, rel)
        if os.path.isfile(p):
            return p
    return None


def build_icon():
    """组装多尺寸 QIcon。返回 None 表示无可用图标。"""
    from PyQt6.QtGui import QIcon

    d = icons_dir()
    if d:
        ic = QIcon()
        added = 0
        for s in ICON_SIZES:
            p = os.path.join(d, "%d.png" % s)
            if os.path.isfile(p):
                ic.addFile(p)
                added += 1
        if added >= 2 and not ic.isNull():
            return ic
        _log().warning("多尺寸图标不足（已加载 %d 张），窗口图标可能无法送达任务栏；"
                       "请执行 python3 tools/make_app_icons.py 重新生成。", added)

    p = legacy_icon_path()
    if p:
        ic = QIcon(p)
        if not ic.isNull():
            _log().warning("使用原始大图作为窗口图标（%s）；Qt6 可能不写入 _NET_WM_ICON，"
                           "任务栏图标可能仍不正确。建议生成多尺寸图标。", os.path.basename(p))
            return ic
    return None


def apply_app_identity(app):
    """设置应用身份并挂上图标。返回所用 QIcon（或 None）。"""
    from PyQt6.QtGui import QGuiApplication

    app.setApplicationName(APP_ID)
    app.setApplicationDisplayName(APP_DISPLAY_NAME)
    app.setOrganizationName(APP_ORG)
    # 关键：桌面文件名。Qt6 会把它写入 _GTK_APPLICATION_ID / _KDE_NET_WM_DESKTOP_FILE，
    # UKUI / GNOME / KDE 据此把窗口关联到 APP_ID.desktop，从而取到自定义图标。
    try:
        QGuiApplication.setDesktopFileName(APP_ID)
    except Exception as e:  # noqa: BLE001
        _log().debug("setDesktopFileName 不可用：%s", e)

    # 同时覆盖 xcb 未打包成包名时的兜底类名（部分环境读 WM_CLASS 的 instance 段）
    try:
        from PyQt6.QtWidgets import QApplication
        QApplication.setApplicationName(APP_ID)
    except Exception:  # noqa: BLE001
        pass

    ic = build_icon()
    if ic is not None:
        app.setWindowIcon(ic)
        _log().info("程序图标：已加载多尺寸图标（%s）", icons_dir() or "回退大图")
    else:
        _log().warning("程序图标：未找到任何图标文件，任务栏将显示默认图标。")
    return ic


def _data_home():
    return os.environ.get("XDG_DATA_HOME") or os.path.join(
        os.path.expanduser("~"), ".local", "share")


def _launcher_command():
    """生成 .desktop 的 Exec：优先用随包启动脚本，保证双击后同终端启动一致。"""
    root = project_root()
    for rel in ("启动汇编程序.sh", "run.sh"):
        p = os.path.join(root, rel)
        if os.path.isfile(p):
            try:
                os.chmod(p, os.stat(p).st_mode | 0o111)
            except Exception:  # noqa: BLE001
                pass
            return '"%s"' % p
    return '"%s" "%s"' % (sys.executable or "python3", os.path.join(root, "app", "main.py"))


def _desktop_text(icon_ref):
    return (
        "[Desktop Entry]\n"
        "Type=Application\n"
        "Version=1.0\n"
        "Name=%s\n"
        "Name[zh_CN]=%s\n"
        "GenericName=公文汇编工具\n"
        "GenericName[zh_CN]=公文汇编工具\n"
        "Comment=%s\n"
        "Comment[zh_CN]=%s\n"
        "Exec=%s\n"
        "Path=%s\n"
        "Icon=%s\n"
        "Terminal=false\n"
        "Categories=Office;WordProcessor;\n"
        "Keywords=公文;汇编;Word;文档;目录;封面;\n"
        # 与 apply_app_identity() 设置的 applicationName 一致，桌面环境据此把窗口关联到本项
        "StartupWMClass=%s\n"
        "StartupNotify=true\n"
        % (APP_NAME, APP_NAME, APP_COMMENT, APP_COMMENT,
           _launcher_command(), project_root(), icon_ref, APP_ID)
    )


def install_desktop_entry():
    """把 .desktop 与图标安装到用户目录（幂等）。返回 .desktop 路径或 None。

    只写用户目录，不需要管理员权限；失败不抛异常，仅记录日志。
    """
    log = _log()
    try:
        data = _data_home()
        apps_dir = os.path.join(data, "applications")
        os.makedirs(apps_dir, exist_ok=True)

        # 1) 安装多尺寸图标到 hicolor 主题目录（供按名称查找 / 主题回退）
        src = icons_dir()
        installed_sizes = 0
        if src:
            for s in ICON_SIZES:
                p = os.path.join(src, "%d.png" % s)
                if not os.path.isfile(p):
                    continue
                dst_dir = os.path.join(data, "icons", "hicolor", "%dx%d" % (s, s), "apps")
                try:
                    os.makedirs(dst_dir, exist_ok=True)
                    shutil.copy2(p, os.path.join(dst_dir, APP_ID + ".png"))
                    installed_sizes += 1
                except Exception:  # noqa: BLE001
                    pass

        # 2) 写 .desktop。Icon 用绝对路径最稳妥：不依赖图标缓存是否已刷新。
        icon_ref = (os.path.join(src, "256.png") if src and
                    os.path.isfile(os.path.join(src, "256.png"))
                    else os.path.join(src, "48.png") if src and
                    os.path.isfile(os.path.join(src, "48.png"))
                    else legacy_icon_path())
        if not icon_ref:
            icon_ref = APP_ID  # 交给主题查找
        desktop_path = os.path.join(apps_dir, APP_ID + ".desktop")
        with open(desktop_path, "w", encoding="utf-8") as f:
            f.write(_desktop_text(icon_ref))
        try:
            os.chmod(desktop_path, 0o755)
        except Exception:  # noqa: BLE001
            pass

        # 3) 刷新图标缓存与桌面数据库（不存在即跳过，失败不影响使用）
        if installed_sizes and shutil.which("gtk-update-icon-cache"):
            subprocess.run(["gtk-update-icon-cache", "-f", "-t",
                            os.path.join(data, "icons", "hicolor")],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=20)
        if shutil.which("update-desktop-database"):
            subprocess.run(["update-desktop-database", apps_dir],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=20)

        log.info("桌面集成：已安装 %s（图标 %d 张，Icon=%s）",
                 desktop_path, installed_sizes, icon_ref)
        return desktop_path
    except Exception as e:  # noqa: BLE001
        log.warning("桌面集成：安装 .desktop 失败（不影响程序使用）：%s", e)
        return None


def log_window_identity(win):
    """把窗口在 X 上的真实身份写进日志（排障用，失败静默）。

    任务栏图标不对时，先看这条日志：

    * `WM_CLASS` 两段都应为 `wenhui-assembler`（instance 段来自 argv[0]，
      由 app/main.py 的 `_adopt_desktop_identity()` 置位；class 段来自
      `applicationName`）；
    * `_NET_WM_ICON` 应非空（空说明 QIcon 尺寸集合不合规）；
    * `_GTK_APPLICATION_ID` / `_KDE_NET_WM_DESKTOP_FILE` 应为 `wenhui-assembler`。

    若都正确而任务栏仍不对，则问题在桌面环境一侧（UKUI panel-daemon 只按
    窗口 app-id / 进程 cmdline 匹配桌面项文件名，不读 StartupWMClass）。
    """
    try:
        info = probe_window_attrs(win.winId())
        if not info:
            return
        _log().info("窗口身份自检：WM_CLASS=%s | _NET_WM_ICON=%s | AppId=%s",
                    info.get("wm_class"), "非空" if info.get("has_icon") else "空/缺失",
                    info.get("gtk_app_id"))
    except Exception as e:  # noqa: BLE001
        _log().debug("窗口身份自检失败：%s", e)


def probe_window_attrs(win_id):
    """诊断用：返回窗口的 WM_CLASS / _NET_WM_ICON 是否非空 / AppId。

    仅用于自检与排障（依赖 xprop，缺失时返回 None）。
    """
    if not shutil.which("xprop"):
        return None
    try:
        out = subprocess.run(["xprop", "-id", str(int(win_id)),
                              "WM_CLASS", "_NET_WM_ICON",
                              "_GTK_APPLICATION_ID", "_KDE_NET_WM_DESKTOP_FILE"],
                             capture_output=True, text=True, timeout=10).stdout
    except Exception:  # noqa: BLE001
        return None
    info = {"raw": out}
    for line in out.splitlines():
        if line.startswith("WM_CLASS"):
            info["wm_class"] = line.split("=", 1)[1].strip()
        elif line.startswith("_NET_WM_ICON"):
            info["has_icon"] = "=" in line and line.split("=", 1)[1].strip() not in ("", "not found.")
        elif line.startswith("_GTK_APPLICATION_ID"):
            info["gtk_app_id"] = line.split("=", 1)[1].strip().strip('"')
    return info
