#!/usr/bin/env bash
# 文档汇编程序 · Linux / 银河麒麟 / 统信UOS 一键启动脚本
# 用法：在终端执行  ./启动汇编程序.sh
#
# 依赖安装策略（离线优先，逐级降级）：
#   1) 本地离线：优先使用随包附带的 wheels/ 目录（pip --no-index --find-links）
#   2) 联网安装：仅当 wheels/ 缺失或与本机不匹配时才尝试（内网环境会失败，属正常）
#
# 运行环境策略（自动选择）：
#   A) .venv   虚拟环境（推荐，与系统环境隔离）
#   B) pylibs/ 目录安装 + PYTHONPATH（当系统无 python3-venv / venv 无 pip 时自动启用）
#
# 注意：本脚本为 LF 换行、UTF-8 无 BOM，请勿用 Windows 记事本打开后另存。

cd "$(dirname "$0")" || exit 1
ROOT="$(pwd)"
WHEELS="$ROOT/wheels"
VENV_DIR="$ROOT/.venv"
LIBS_DIR="$ROOT/pylibs"
REQUIRED="python-docx lxml PyQt6"
OPTIONAL="PyMuPDF"

log() { printf '%s\n' "$*"; }

# ---------------------------------------------------------------- 1. 探测 Python
find_python() {
    local c
    for c in python3 python3.13 python3.12 python3.11 python3.10 python3.9 python3.8 python; do
        if command -v "$c" >/dev/null 2>&1; then
            if "$c" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 8) else 1)' 2>/dev/null; then
                printf '%s' "$c"
                return 0
            fi
        fi
    done
    return 1
}

if ! PY="$(find_python)"; then
    log "错误：未检测到 Python 3.8 及以上版本。"
    log "请先安装："
    log "  银河麒麟/统信UOS/Debian/Ubuntu:  sudo apt install python3 python3-venv python3-pip"
    log "  CentOS/RHEL/Fedora:               sudo yum install python3 python3-pip"
    exit 1
fi

PYVER="$("$PY" -c 'import sys;print("%d.%d" % sys.version_info[:2])')"
# 架构名归一化：不同环境会返回 AMD64 / x86_64 / arm64 / aarch64，统一成 wheel 平台标签的写法
_ARCH_PY='import platform
m = platform.machine().lower()
alias = {"amd64": "x86_64", "x64": "x86_64", "i386": "x86_64", "i686": "x86_64",
         "arm64": "aarch64", "armv8l": "aarch64", "armv7l": "armv7l"}
print(alias.get(m, m))'
ARCH="$("$PY" -c "$_ARCH_PY")"
log "=============================================================="
log " 文档汇编程序"
log " Python $PYVER   架构 $ARCH"
log " 目录 $ROOT"
log "=============================================================="

# 校验离线包架构是否与本机匹配
if [ -d "$WHEELS" ]; then
    WHL_ARCH="$(ls "$WHEELS" 2>/dev/null | grep -o 'aarch64\|x86_64' | head -1)"
    if [ -n "$WHL_ARCH" ] && [ "$WHL_ARCH" != "$ARCH" ]; then
        log "警告：wheels/ 内的离线包为 $WHL_ARCH 架构，与本机 $ARCH 不匹配，离线安装会失败。"
        log "      请联系分发方索取 $ARCH 版本的离线包，或改用联网安装。"
    fi
fi

# 校验 glibc 版本：PyQt6 官方 wheel 仅提供 manylinux_2_28，要求 glibc >= 2.28
GLIBC_OK=1
GLIBC="$(ldd --version 2>/dev/null | head -1 | grep -oE '[0-9]+\.[0-9]+' | head -1)"
if [ -n "$GLIBC" ]; then
    G_MAJ="${GLIBC%%.*}"
    G_MIN="${GLIBC##*.}"
    [ -n "$G_MIN" ] || G_MIN=0
    if [ "$G_MAJ" -lt 2 ] 2>/dev/null || { [ "$G_MAJ" -eq 2 ] && [ "$G_MIN" -lt 28 ]; }; then
        GLIBC_OK=0
        log "警告：本机 glibc 为 $GLIBC，低于 PyQt6 官方 wheel 要求的 2.28。"
        log "      离线安装 PyQt6 很可能会失败。建议改用系统软件源安装："
        log "        sudo apt install python3-pyqt6 python3-docx python3-lxml"
        log "      装好后重新执行本脚本即可（脚本会自动识别系统已装依赖，不再联网）。"
    fi
fi

# 校验离线包是否包含启动所需的所有 wheel
# （避免"装到一半才发现缺 PyQt6"，提前给明确提示）
check_wheels_present() {
    [ -d "$WHEELS" ] || return 0
    local need prefix miss=0
    for need in python-docx lxml PyQt6; do
        case "$need" in
            python-docx) prefix="python_docx" ;;
            lxml)        prefix="lxml" ;;
            PyQt6)       prefix="PyQt6-" ;;
        esac
        if ! ls "$WHEELS"/${prefix}*.whl >/dev/null 2>&1; then
            log "警告：wheels/ 内缺少 $need 的离线包（${prefix}*.whl），离线安装会失败。"
            miss=1
        fi
    done
    if [ "$miss" = "1" ]; then
        log "      请联系分发方确认该离线包针对 $ARCH 架构完整打包，或改用联网安装。"
    fi
}
check_wheels_present

# ---------------------------------------------------------------- 输入法支持
# Qt6 界面输入中文依赖对应输入法模块（QT_IM_MODULE）。麒麟/统信常用 fcitx / fcitx5 / ibus。
# 未设置或缺少 Qt6 输入法插件时，界面输入框会"无法切换输入法 / 未识别输入窗口"。
# 此处自动探测并导出环境变量；若仍无法输入中文，需补装对应 Qt6 插件（见 国产系统运行说明.md）：
#   fcitx/fcitx5:  sudo apt install fcitx-frontend-qt6   （fcitx5 用 fcitx5-frontend-qt6）
#   ibus:          sudo apt install ibus-qt6              （或 libqt6-ibus-platforminputcontext）
detect_im() {
    # fcitx5 与 fcitx4 注册的 Qt 模块名不同（fcitx5 / fcitx），必须区分，否则 Qt 找不到对应插件
    if pgrep -x fcitx5 >/dev/null 2>&1 || command -v fcitx5 >/dev/null 2>&1; then
        printf 'fcitx5'
    elif pgrep -x fcitx >/dev/null 2>&1 || command -v fcitx >/dev/null 2>&1; then
        printf 'fcitx'
    elif pgrep -x ibus-daemon >/dev/null 2>&1 || command -v ibus-daemon >/dev/null 2>&1; then
        printf 'ibus'
    else
        printf ''
    fi
}
IM_MODULE="$(detect_im)"
if [ -n "$IM_MODULE" ]; then
    export QT_IM_MODULE="$IM_MODULE"
    export GTK_IM_MODULE="$IM_MODULE"
    export XMODIFIERS="@im=$IM_MODULE"
    log "输入法：$IM_MODULE（已设置 QT_IM_MODULE=$IM_MODULE）"
else
    log "提示：未检测到 fcitx/ibus 输入法，中文输入可能不可用；可手动 export QT_IM_MODULE=fcitx 后重试。"
fi

# ---------------------------------------------------------------- 输入法插件接入
# pip 安装的 PyQt6 自带一套 Qt6，其插件目录在 <site-packages>/PyQt6/Qt6/plugins 下，
# 默认不会去搜索系统的 /usr/lib/qt6/plugins。因此即使系统装了 fcitx-frontend-qt6 等
# 输入法插件，PyQt6 也加载不到、界面仍无法输入中文。
# 此函数把系统里已装的 Qt6 输入法插件（fcitx/fcitx5/ibus）复制进 PyQt6 自己的插件目录，
# 使 QT_IM_MODULE 指定的输入法真正生效。
# 用法：link_im_plugin <python> [PYTHONPATH]  （PYTHONPATH 用于 pylibs 目录安装方案）
link_im_plugin() {
    local py="$1" pp="$2"
    PYTHONPATH="$pp" "$py" - 2>/dev/null <<'PYEOF'
import os, shutil, sys
try:
    import PyQt6
except ImportError:
    sys.exit(0)
pkg = os.path.dirname(os.path.abspath(PyQt6.__file__))
plug = os.path.join(pkg, "Qt6", "plugins")
dst = os.path.join(plug, "platforminputcontexts")
if not os.path.isdir(plug):
    sys.exit(0)
os.makedirs(dst, exist_ok=True)
found = 0
for d in ("/usr/lib/qt6/plugins/platforminputcontexts",
          "/usr/lib/aarch64-linux-gnu/qt6/plugins/platforminputcontexts",
          "/usr/lib/x86_64-linux-gnu/qt6/plugins/platforminputcontexts",
          "/usr/lib64/qt6/plugins/platforminputcontexts"):
    if not os.path.isdir(d):
        continue
    for fn in os.listdir(d):
        if not fn.endswith(".so"):
            continue
        t = os.path.join(dst, fn)
        if not os.path.exists(t):
            try:
                shutil.copy2(os.path.join(d, fn), t)
                print("输入法插件：已接入", fn)
                found = 1
            except Exception:
                pass
if not found:
    print("提示：未找到系统 Qt6 输入法插件，中文输入可能仍不可用；可执行 sudo apt install fcitx5-frontend-qt6")
PYEOF
}

# ---------------------------------------------------------------- 字体安装（免 root）
# 公文排版需 仿宋_GB2312 / 方正小标宋 / 楷体_GB2312 / 黑体 / 宋体 等字体，麒麟/统信常缺失。
# 把随包字体复制到用户字体目录并刷新字体缓存，使 python-docx 排版与 LibreOffice 转 PDF
# 都能正确渲染公文版式（LibreOffice 通过 fontconfig 发现这些字体）。
install_bundled_fonts() {
    local font_dir="$ROOT/字体包"
    [ -d "$font_dir" ] || font_dir="$ROOT/app/fonts_bundled"
    [ -d "$font_dir" ] || return 0
    local target="${XDG_DATA_HOME:-$HOME/.local/share}/fonts"
    mkdir -p "$target" 2>/dev/null || return 0
    local n=0 f bn
    for f in "$font_dir"/*.ttf "$font_dir"/*.ttc "$font_dir"/*.otf; do
        [ -f "$f" ] || continue
        bn="$(basename "$f")"
        if [ ! -f "$target/$bn" ]; then
            cp "$f" "$target/$bn" 2>/dev/null && n=$((n+1))
        fi
    done
    if [ "$n" -gt 0 ]; then
        if command -v fc-cache >/dev/null 2>&1; then
            fc-cache -f "$target" >/dev/null 2>&1
        fi
        log "字体：已安装 $n 个随包字体到 $target（免 root）"
    fi
}
install_bundled_fonts

# ---------------------------------------------------------------- LibreOffice 就位（免 root）
# 随包便携 LibreOffice 解压后需确保可执行（部分解压工具会丢失可执行位）。
fix_soffice_exec() {
    local lo="$ROOT/app/libreoffice/opt"
    [ -d "$lo" ] || return 0
    find "$lo" -type d -name program -exec chmod -R a+x {} \; 2>/dev/null
}
fix_soffice_exec

# ---------------------------------------------------------------- 2. 依赖检测
deps_ok() {
    "$1" -c 'import docx, lxml; from PyQt6.QtWidgets import QApplication' >/dev/null 2>&1
}

pymupdf_ok() {
    "$1" -c 'import pymupdf' >/dev/null 2>&1 || "$1" -c 'import fitz' >/dev/null 2>&1
}

# ---------------------------------------------------------------- 3. pip 调用封装
# 若环境无 pip，用 wheels/ 内的 pip wheel 以 PYTHONPATH 方式临时调用（wheel 即 zip，可直接导入）。
# wheels/ 内可能同时存在多个 pip 版本（如 24.0 与 25.0.1），逐个探测，
# 取"最后一个可用"——glob 按文件名升序，因此最后一个即版本号最高的可用 pip。
PIP_WHL=""
select_pip_whl() {
    local py="$1" f best=""
    [ -d "$WHEELS" ] || return 1
    for f in "$WHEELS"/pip-*.whl; do
        [ -f "$f" ] || continue
        if PYTHONPATH="$f" "$py" -m pip --version >/dev/null 2>&1; then
            best="$f"
        fi
    done
    [ -n "$best" ] || return 1
    printf '%s' "$best"
}

have_pip() { "$1" -m pip --version >/dev/null 2>&1; }

run_pip() {
    local py="$1"; shift
    if [ -n "$PIP_WHL" ]; then
        # 优先使用随包自带的新版 pip（如 pip-25.0.1）。
        # 系统自带 pip 可能过旧（如 20.0.2），不认识 manylinux_2_28 等新平台标签，
        # 会导致 PyQt6 等离线 wheel 被判定为"标签不匹配"而无法安装。
        PYTHONPATH="$PIP_WHL" "$py" -m pip "$@"
    elif have_pip "$py"; then
        "$py" -m pip "$@"
    else
        return 1
    fi
}

bootstrap_pip() {
    local py="$1"
    have_pip "$py" && return 0
    log "  环境缺少 pip，尝试本地引导 ..."
    if [ -n "$PIP_WHL" ]; then
        # 只装 pip 本体：setuptools 非必需，且其 Requires-Python 可能在老版本 Python 上不满足
        run_pip "$py" install --no-index --find-links="$WHEELS" pip >/dev/null 2>&1
        have_pip "$py" && return 0
    fi
    "$py" -m ensurepip --default-pip >/dev/null 2>&1
    have_pip "$py"
}

# 选出当前 Python 可用的离线 pip 引导包。
# 必须在 select_pip_whl 定义之后调用：bash 是顺序执行，提前调用会 command not found。
# venv 与系统 Python 版本一致，探测一次即可。
PIP_WHL="$(select_pip_whl "$PY")" || PIP_WHL=""
if [ -n "$PIP_WHL" ]; then
    log "离线 pip 引导包：$(basename "$PIP_WHL")"
fi

# ---------------------------------------------------------------- 4. 安装依赖
install_required() {
    local py="$1"; shift
    local extra="$*"
    if [ -d "$WHEELS" ]; then
        log "  离线安装（wheels/ 本地包）..."
        if run_pip "$py" install $extra --no-index --find-links="$WHEELS" $REQUIRED; then
            return 0
        fi
        log "  离线安装未成功，尝试联网安装（内网环境会失败，属正常现象）..."
    fi
    run_pip "$py" install $extra $REQUIRED
}

install_optional() {
    local py="$1"; shift
    local extra="$*"
    pymupdf_ok "$py" && return 0
    if [ -d "$WHEELS" ]; then
        run_pip "$py" install $extra --no-index --find-links="$WHEELS" $OPTIONAL >/dev/null 2>&1 && return 0
    fi
    run_pip "$py" install $extra $OPTIONAL >/dev/null 2>&1 && return 0
    return 1
}

exec_run() {
    # 启动前把系统 Qt6 输入法插件接入 PyQt6（否则界面无法输入中文）
    link_im_plugin "$1" "$2"
    if [ -n "$2" ]; then
        exec env PYTHONPATH="$2" "$1" app/main.py
    else
        exec "$1" app/main.py
    fi
}

# ---------------------------------------------------------------- 5. 主流程
mkdir -p "$ROOT/logs" 2>/dev/null || true

# 方案 0：系统 Python 已具备依赖
if deps_ok "$PY"; then
    log "系统 Python 已具备全部依赖，直接启动。"
    exec_run "$PY" ""
fi

# 方案 A：虚拟环境
VENV_OK=0
if [ -x "$VENV_DIR/bin/python" ]; then
    VENV_OK=1
else
    log "正在创建虚拟环境 .venv ..."
    if "$PY" -m venv "$VENV_DIR" >/dev/null 2>&1 && [ -x "$VENV_DIR/bin/python" ]; then
        VENV_OK=1
    else
        log "  虚拟环境创建失败（可能未安装 python3-venv），将改用目录安装方案。"
        rm -rf "$VENV_DIR" 2>/dev/null || true
    fi
fi

if [ "$VENV_OK" = "1" ]; then
    VPY="$VENV_DIR/bin/python"
    if deps_ok "$VPY"; then
        log "虚拟环境依赖已就绪，启动程序。"
        exec_run "$VPY" ""
    fi
    log "正在为虚拟环境准备依赖 ..."
    bootstrap_pip "$VPY" || true
    if install_required "$VPY"; then
        install_optional "$VPY" || log "  提示：PyMuPDF 未安装，预览功能不可用（不影响汇编导出）。"
        if deps_ok "$VPY"; then
            log "依赖就绪，启动程序。"
            exec_run "$VPY" ""
        fi
        log "  虚拟环境方案安装后仍无法导入 PyQt6，改用目录安装方案。"
    else
        log "  虚拟环境方案安装失败，改用目录安装方案。"
    fi
fi

# 方案 B：目录安装 + PYTHONPATH（不改动系统环境，无需管理员权限）
log "正在将依赖安装到 pylibs/ 目录 ..."
if install_required "$PY" "--target=$LIBS_DIR"; then
    install_optional "$PY" "--target=$LIBS_DIR" || log "  提示：PyMuPDF 未安装，预览功能不可用（不影响汇编导出）。"
    if PYTHONPATH="$LIBS_DIR" "$PY" -c 'import docx, lxml; from PyQt6.QtWidgets import QApplication' >/dev/null 2>&1; then
        log "依赖就绪，启动程序。"
        exec_run "$PY" "$LIBS_DIR"
    fi
fi

# ---------------------------------------------------------------- 6. 诊断信息
log ""
log "=============================================================="
log " 启动失败：依赖未能安装完成"
log "=============================================================="
log "请按顺序排查："
log ""
log "一、离线包是否与本机匹配"
log "    wheels/ 目录应存在，且其中的 wheel 需为 $ARCH 架构、支持 Python $PYVER。"
log "    （随包离线包覆盖 Python 3.8 ~ 3.12；若本机为更高版本，请用系统源安装后再运行本脚本）"
log "    查看本机架构：uname -m      查看 Python 版本：python3 -V"
log ""
log "二、缺少 PyPI 依赖且无法联网时，改用系统源安装："
log "    银河麒麟/统信UOS/Debian/Ubuntu:"
log "      sudo apt install python3-docx python3-lxml python3-pyqt6"
log "    CentOS/RHEL/Fedora:"
log "      sudo yum install python3-docx python3-lxml python3-qt6"
log "    安装后重新执行本脚本即可（脚本会自动识别系统已装依赖）。"
log ""
log "三、PyQt6 已安装但无法导入，通常是缺少系统图形库："
log "    银河麒麟/统信UOS/Debian/Ubuntu:"
log "      sudo apt install libgl1 libegl1 libxkbcommon-x11-0 libdbus-1-3 \\"
log "                       libxcb-icccm4 libxcb-image0 libxcb-keysyms1 \\"
log "                       libxcb-render-util0 libxcb-shape0 libxcb-xinerama0 \\"
log "                       libxcb-cursor0 libxcb-randr0 libxcb-xfixes0"
log ""
log "四、可将本目录下的 logs/ 内最新日志文件提供给技术支持。"
log "=============================================================="
exit 1
