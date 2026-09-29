# -*- coding: utf-8 -*-
"""
converter.py —— 读取与转换模块（转换引擎按平台选择）
把 .doc / .wps / .docx 统一转为 .docx（临时文件），交给 assembler 处理。
预览时再把成品 .docx 转 PDF。

探测方式：
- Windows：优先 WPS（COM 自动化；个人版/专业版均可），兜底 Microsoft Word COM。
- Linux 国产系统（麒麟/统信）：优先 LibreOffice 无头转换
  （`soffice --headless --convert-to`）；WPS on Linux 不支持命令行转换，
  仅在 LibreOffice 缺失时兜底尝试 WPS 命令行。

提供 detect_converter() / wps_can_convert() 做引擎探测与能力自检。
"""
import os
import glob
import re
import shutil
import subprocess
import tempfile
import time
import platform
import zipfile
from xml.etree import ElementTree as ET

try:
    import winreg
except ImportError:
    winreg = None

# 程序目录（app/），用于定位随程序集成的便携 LibreOffice（国产系统离线包内，免 root）
_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _is_exe(p):
    return bool(p) and os.path.isfile(p) and os.access(p, os.X_OK)


def _reg_app_path(exe_name):
    """Windows：从 App Paths 取 exe 全路径（HKLM + HKCU，覆盖本机/当前用户安装）。"""
    if winreg is None:
        return None
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for sub in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"):
            try:
                key = winreg.OpenKey(root, sub + "\\" + exe_name)
                val, _ = winreg.QueryValueEx(key, None)
                if val and _is_exe(val):
                    return val
            except OSError:
                pass
    return None


def _glob_first(patterns):
    for pat in patterns:
        for m in glob.glob(pat):
            if _is_exe(m):
                return m
    return None


def _find_wps():
    """探测 WPS 可执行文件。Windows 优先注册表；Linux 麒麟/统信通常自带。"""
    if platform.system() == "Windows":
        c = _reg_app_path("wps.exe")
        if c:
            return c
        drives = ["C", "D", "E", "F"]
        pats = []
        for d in drives:
            for pf in ("Program Files", "Program Files (x86)"):
                pats.append(rf"{d}:\{pf}\WPS Office\office6\wps.exe")
                pats.append(rf"{d}:\{pf}\WPS Office\*\office6\wps.exe")
                pats.append(rf"{d}:\{pf}\Kingsoft\WPS Office\office6\wps.exe")
                pats.append(rf"{d}:\{pf}\Kingsoft\WPS Office\*\office6\wps.exe")
        return _glob_first(pats)
    else:
        for name in ("wps",):
            f = shutil.which(name)
            if f and _is_exe(f):
                return f
        for p in ("/usr/bin/wps",
                  "/opt/kingsoft/wps-office/office6/wps",
                  "/usr/lib/wps-office/office6/wps"):
            if _is_exe(p):
                return p
    return None


def detect_engine():
    """返回 WPS 可执行文件路径，未检测到返回 None。"""
    return _find_wps()


def _find_soffice():
    """探测 LibreOffice 可执行文件（无头转换，Linux/Windows 通用）。

    国产系统（麒麟/统信）WPS 不支持命令行转换，LibreOffice 的
    `soffice --headless --convert-to` 是可靠的无头转换方案。

    探测顺序：① 系统 PATH 与常见安装路径 → ② 随程序集成的便携 LibreOffice
    （国产系统离线包内 app/libreoffice/opt/libreoffice*/program/soffice，免 root）。
    """
    for name in ("soffice", "libreoffice", "soffice.bin"):
        f = shutil.which(name)
        if f and _is_exe(f):
            return f
    for p in ("/usr/bin/soffice", "/usr/bin/libreoffice",
              "/opt/libreoffice/program/soffice",
              "/usr/lib/libreoffice/program/soffice",
              "C:/Program Files/LibreOffice/program/soffice.exe"):
        if _is_exe(p):
            return p
    # 随程序集成的便携 LibreOffice（deb 包解出后：app/libreoffice/opt/libreoffice<版本>/program/soffice）
    bundled_root = os.path.join(_APP_DIR, "libreoffice")
    for pat in ("opt/libreoffice*/program/soffice",
                "opt/libreoffice*/program/soffice.bin",
                "program/soffice", "program/soffice.bin"):
        for m in glob.glob(os.path.join(bundled_root, pat)):
            if _is_exe(m):
                return m
    return None


def detect_converter():
    """返回当前平台可用的转换引擎 (name, path)。

    Windows: WPS（COM 自动化）；Linux 国产系统: LibreOffice（无头命令行）。
    供界面状态提示与能力自检使用。
    """
    if platform.system() == "Windows":
        exe = detect_engine()
        return ("WPS", exe) if exe else (None, None)
    soffice = _find_soffice()
    if soffice:
        return ("LibreOffice", soffice)
    wps = detect_engine()
    if wps:
        return ("WPS", wps)
    return (None, None)


def _kill_proc(exe):
    name = os.path.basename(exe)
    try:
        if platform.system() == "Windows":
            subprocess.run(["taskkill", "/F", "/IM", name],
                           capture_output=True, timeout=10)
        else:
            subprocess.run(["pkill", "-f", name], capture_output=True, timeout=10)
    except Exception:
        pass


def _convert_with_com(exe, src, out_dir, fmt):
    """Windows：WPS COM 自动化转换（个人版/专业版均可）。返回输出路径。

    通过 COM 接口「打开->另存为->关闭」完成 docx/pdf 输出，
    绕开 WPS 个人版不支持无头命令行转换的问题。
    连续转换时 COM 实例偶发失败，故内置一次重试。
    """
    try:
        import win32com.client as win32
    except Exception:
        raise RuntimeError("未安装 pywin32，无法使用 WPS COM 转换")
    work = tempfile.mkdtemp(prefix="com_")
    ext = os.path.splitext(src)[1].lower() or ".docx"
    safe_src = os.path.join(work, "src" + ext)
    try:
        shutil.copy2(src, safe_src)
    except Exception as e:
        raise RuntimeError(f"准备转换源文件失败：{e}")
    fmt_num = 16 if fmt == "docx" else 17  # wdFormatXMLDocument / wdFormatPDF
    last = None
    for attempt in (1, 2):
        app = None
        try:
            app = win32.Dispatch("KWps.Application")
            try:
                app.Visible = False
            except Exception:
                pass
            doc = app.Documents.Open(safe_src)
            out = os.path.join(work, "src." + fmt)
            try:
                doc.SaveAs2(out, fmt_num)
            finally:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if os.path.isfile(out) and os.path.getsize(out) > 0:
                return out
            last = RuntimeError("COM 未生成有效输出")
        except Exception as e:
            last = RuntimeError("WPS COM 转换失败：" + str(e)[:200])
        finally:
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
        if attempt == 1:
            time.sleep(1.5)  # 连续转换时给 COM 实例退出留时间
    raise last


def _convert_with(exe, src, out_dir, fmt, timeout=None):
    """用 WPS 把 src 转为 .fmt（docx / pdf）。失败抛异常。

    timeout：子进程超时秒数，None 表示用默认 120 秒（供能力自检收紧超时）。

    注意：WPS 命令行对含非 ASCII（如中文）的路径常加载失败，
    故先把源文件复制到 ASCII 临时名再转换，输出也落在该 ASCII 临时目录。
    """
    # 复制到 ASCII 临时工作目录，规避中文路径问题
    # （同样优先落在输出磁盘，避免 /tmp 为小容量 tmpfs 时写不出结果）
    try:
        os.makedirs(out_dir, exist_ok=True)
        work = tempfile.mkdtemp(prefix="conv_", dir=out_dir)
    except Exception:
        work = tempfile.mkdtemp(prefix="conv_")
    ext = os.path.splitext(src)[1].lower() or ".docx"
    safe_src = os.path.join(work, "src" + ext)
    try:
        shutil.copy2(src, safe_src)
    except Exception as e:
        raise RuntimeError(f"准备转换源文件失败：{e}")
    safe_out = work
    cmd = [exe, "--headless", "--convert-to", fmt, "--outdir", safe_out, safe_src]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout or 120)
    except subprocess.TimeoutExpired:
        _kill_proc(exe)
        raise RuntimeError("WPS 转换超时（可能不支持无头命令行转换，请使用 WPS 专业版/政府版）")
    if r.returncode != 0:
        raise RuntimeError(
            "WPS 转换失败：" + (r.stderr or r.stdout or "返回非零") +
            "\n提示：WPS 个人版通常不支持命令行转换，请使用 WPS 专业版/政府版。")
    # 推断输出文件名（基于安全名 src）
    target = os.path.join(safe_out, "src." + fmt)
    if not os.path.isfile(target):
        for f in os.listdir(safe_out):
            if f.lower().endswith("." + fmt) and "src" in f:
                target = os.path.join(safe_out, f)
                break
    if not os.path.isfile(target):
        raise RuntimeError(f"转换后未找到 .{fmt} 文件")
    return target


def _convert_with_word_com(src, out_dir, fmt):
    """用 Microsoft Word COM 转换（WPS 不可用时的兜底）。

    wdFormat 常量与 WPS COM 一致：docx=16 (wdFormatXMLDocument)、pdf=17 (wdFormatPDF)。
    同样把源文件复制到 ASCII 临时目录，规避中文路径问题。
    """
    try:
        import win32com.client as win32
    except Exception:
        raise RuntimeError("未安装 pywin32，无法使用 Word COM 转换")
    work = tempfile.mkdtemp(prefix="wordcom_")
    ext = os.path.splitext(src)[1].lower() or ".docx"
    safe_src = os.path.join(work, "src" + ext)
    try:
        shutil.copy2(src, safe_src)
    except Exception as e:
        raise RuntimeError(f"准备转换源文件失败：{e}")
    fmt_num = 16 if fmt == "docx" else 17
    last = None
    for attempt in (1, 2):
        app = None
        try:
            app = win32.Dispatch("Word.Application")
            try:
                app.Visible = False
                app.DisplayAlerts = 0
            except Exception:
                pass
            doc = app.Documents.Open(safe_src, ReadOnly=True)
            out = os.path.join(work, "src." + fmt)
            try:
                doc.SaveAs2(out, fmt_num)
            finally:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if os.path.isfile(out) and os.path.getsize(out) > 0:
                return out
            last = RuntimeError("Word COM 未生成有效输出")
        except Exception as e:
            last = RuntimeError("Word COM 转换失败：" + str(e)[:200])
        finally:
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
        if attempt == 1:
            time.sleep(1.5)  # 给 Word COM 实例退出留时间（残留实例会导致后续调用失败）
    raise last


def _convert_with_soffice(exe, src, out_dir, fmt, timeout=None):
    """用 LibreOffice 无头转换（Linux/Windows 通用，国产系统推荐）。

    timeout：子进程超时秒数，None 表示用默认 180 秒（供能力自检收紧超时）。

    通过 `soffice --headless --convert-to <fmt> --outdir <out> <src>` 完成转换。
    使用独立 UserInstallation 配置目录，避免与系统已运行的 LibreOffice 实例冲突。

    环境处理：随包便携 LibreOffice 把第三方依赖库与主程序放在同一目录
    （program/），其二进制 RUNPATH 为 $ORIGIN。RUNPATH 对“间接依赖”不生效
    （如 libclucene-core 找 libclucene-shared 时不继承调用者的 RUNPATH），
    故这里显式把该目录加入 LD_LIBRARY_PATH，避免 "cannot open shared object file"。
    """
    # 临时工作目录优先放在输出目录所在的磁盘：汇编产物可达上百页并含图片，
    # 转出的 PDF 可能数十 MB，而部分国产系统 /tmp 是容量很小的 tmpfs，
    # 会导致 LibreOffice 写 PDF 失败（返回码仍为 0，只会静默不产出文件）。
    try:
        os.makedirs(out_dir, exist_ok=True)
        work = tempfile.mkdtemp(prefix="lo_", dir=out_dir)
    except Exception:
        work = tempfile.mkdtemp(prefix="lo_")
    ext = os.path.splitext(src)[1].lower() or ".docx"
    safe_src = os.path.join(work, "src" + ext)
    try:
        shutil.copy2(src, safe_src)
    except Exception as e:
        raise RuntimeError(f"准备转换源文件失败：{e}")
    user_install = "file://" + os.path.join(work, "lo_profile").replace("\\", "/")
    cmd = [exe, "--headless", "--norestore", "--invisible", "--nodefault",
           "-env:UserInstallation=" + user_install,
           "--convert-to", fmt, "--outdir", work, safe_src]
    env = dict(os.environ)
    exe_dir = os.path.dirname(os.path.abspath(exe))
    if exe_dir:
        prev = env.get("LD_LIBRARY_PATH", "")
        env["LD_LIBRARY_PATH"] = exe_dir + (":" + prev if prev else "")
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           timeout=timeout or 180, env=env)
    except subprocess.TimeoutExpired:
        _kill_proc(exe)
        raise RuntimeError("LibreOffice 转换超时（文档过大或含不支持对象）")
    if r.returncode != 0:
        raise RuntimeError(
            "LibreOffice 转换失败：" + (r.stderr or r.stdout or "返回非零")[:300])
    target = os.path.join(work, "src." + fmt)
    if not os.path.isfile(target):
        for f in os.listdir(work):
            if f.lower().endswith("." + fmt):
                target = os.path.join(work, f)
                break
    if not os.path.isfile(target) or os.path.getsize(target) == 0:
        # LibreOffice 在"能读源文件但写不出结果"时返回码仍为 0（例如输出磁盘
        # 空间不足、字体缺失致渲染中断），此时 stdout/stderr 是唯一线索，
        # 必须一并带出，否则用户只看到"未找到 .pdf"而无法定位。
        detail = ((r.stderr or "") + (r.stdout or "")).strip()
        raise RuntimeError(
            f"LibreOffice 转换后未找到有效的 .{fmt} 文件"
            + (f"（引擎输出：{detail[:400]}）" if detail else
               "（引擎无输出；请检查输出磁盘剩余空间与字体是否完整）"))
    return target


def _run_convert(src, out_dir, fmt, timeout=None):
    """转换策略（按平台选择引擎）：

    - Windows：WPS COM → WPS 命令行 → Microsoft Word COM（兜底）。
    - Linux（国产系统）：优先 LibreOffice 无头转换；WPS on Linux 不支持命令行
      转换，仅在 LibreOffice 缺失时兜底尝试 WPS 命令行。

    timeout：把子进程超时一路透传给命令行引擎（None 用各引擎默认值）。
    COM 引擎没有子进程可超时，其自身带重试，不受此参数影响。
    """
    errors = []
    if platform.system() == "Windows":
        exe = detect_engine()
        if exe:
            try:
                return _convert_with_com(exe, src, out_dir, fmt)
            except Exception as e:
                errors.append("WPS COM: " + str(e))
            try:
                return _convert_with(exe, src, out_dir, fmt, timeout=timeout)
            except Exception as e:
                errors.append("WPS 命令行: " + str(e))
        try:
            return _convert_with_word_com(src, out_dir, fmt)
        except Exception as e:
            errors.append("Word COM: " + str(e))
    else:
        soffice = _find_soffice()
        if soffice:
            try:
                return _convert_with_soffice(soffice, src, out_dir, fmt,
                                             timeout=timeout)
            except Exception as e:
                msg = str(e)
                # 便携版与发行版二进制常见"glibc 版本不足"故障：进程起不来，
                # 报错只有动态链接器的 GLIBC_x.y not found，容易被当成"转换失败"。
                if "GLIBC_" in msg and "not found" in msg:
                    msg += ("\n提示：该 LibreOffice 与本机 glibc 不兼容（版本过低）。"
                            "请安装系统版，程序会优先使用它：\n"
                            "  sudo apt install libreoffice-writer")
                errors.append("LibreOffice: " + msg)
        else:
            errors.append("LibreOffice: 未检测到（请安装 LibreOffice，"
                          "或确认便携版已随离线包部署）")
        wps = detect_engine()
        if wps:
            try:
                return _convert_with(wps, src, out_dir, fmt, timeout=timeout)
            except Exception as e:
                errors.append("WPS 命令行: " + str(e))
    raise RuntimeError("转换失败（已尝试 " + " / ".join(errors) + "）：\n"
                       + "\n".join(errors))


def _docx_needs_normalize(src_path):
    """判断 .docx 是否能被 python-docx 正常打开。

    python-docx 仅接受标准正文文档（content type =
    wordprocessingml.document.main+xml）。宏启用文档(.docm，content type
    含 macroEnabled) 或模板(.dotx) 等会被其拒绝（抛 "not a Word file"），
    这类文件需规范化（剥离宏）后再汇编。
    """
    try:
        from docx import Document
        Document(src_path)
        return False
    except Exception:
        return True


# OOXML 标准正文文档 content type（python-docx 唯一接受的值）
_CT_DOC_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _normalize_docx_package(src_path, out_dir):
    """轻量规范化：剥离宏（vbaProject），把主文档 Part 的 ContentType 改回标准 docx，
    并清除对 vbaProject 的引用关系；全部文本与排版原样保留。不依赖 WPS。

    适用于 python-docx 因 "macroEnabled" 等拒绝打开的 .docx（多为 .docm
    被误存为 .docx 扩展名）。返回规范化后的新 .docx 路径。
    """
    work = tempfile.mkdtemp(prefix="norm_")
    out = os.path.join(work, "normalized.docx")
    try:
        with zipfile.ZipFile(src_path, "r") as zin:
            parts = {n: zin.read(n) for n in zin.namelist()}
    except Exception as e:
        raise RuntimeError("读取源 docx 失败（可能不是有效压缩包）：" + str(e)[:200])

    # 1) 修正 [Content_Types].xml：主文档 Override 改回标准类型，删除 macroEnabled 项
    root = None
    changed = False
    if "[Content_Types].xml" in parts:
        try:
            root = ET.fromstring(parts["[Content_Types].xml"])
        except ET.ParseError as e:
            raise RuntimeError("解析 [Content_Types].xml 失败：" + str(e))
        for el in list(root):
            tag = el.tag.split("}")[-1]
            ct = el.get("ContentType", "")
            if tag == "Override" and el.get("PartName") == "/word/document.xml":
                if ct != _CT_DOC_MAIN:
                    el.set("ContentType", _CT_DOC_MAIN)
                    changed = True
                continue  # 主文档 Part 已处理，勿按 macroEnabled 误删
            if "macroEnabled" in ct:
                root.remove(el)
                changed = True
            elif tag == "Default" and el.get("Extension", "").lower() == "bin" \
                    and "vbaProject" in ct:
                root.remove(el)
                changed = True
        if changed:
            parts["[Content_Types].xml"] = ET.tostring(
                root, encoding="UTF-8", xml_declaration=True)

    # 2) 删除 vbaProject.bin 部件，并清除其对应的 Content-Types Override
    vb = "word/vbaProject.bin"
    if vb in parts and root is not None:
        del parts[vb]
        for el in list(root):
            if el.tag.split("}")[-1] == "Override" \
                    and (el.get("PartName", "").endswith("vbaProject.bin")):
                root.remove(el)
                changed = True

    # 3) 清除 document.xml.rels 中对 vbaProject 的引用，避免悬空关系
    rels = "word/_rels/document.xml.rels"
    if rels in parts:
        try:
            rroot = ET.fromstring(parts[rels])
            removed = False
            for rel in list(rroot):
                if rel.tag.split("}")[-1] == "Relationship" \
                        and rel.get("Target", "").endswith("vbaProject.bin"):
                    rroot.remove(rel)
                    removed = True
            if removed:
                parts[rels] = ET.tostring(
                    rroot, encoding="UTF-8", xml_declaration=True)
        except ET.ParseError:
            pass

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in parts.items():
            zout.writestr(n, b)

    # 校验：规范化后必须能被 python-docx 打开
    try:
        from docx import Document
        Document(out)
    except Exception as e:
        raise RuntimeError("规范化后仍无法打开：" + str(e)[:200])
    return out


# python-docx 1.1.x 认识的 w:jc 取值（WD_PARAGRAPH_ALIGNMENT）
_JC_KNOWN = {"left", "right", "center", "both", "distribute"}
# 等价改写：ECMA-376 第 4 版（Word 2013+ / LibreOffice 6.4+ 导出）用 start/end
_JC_ALIAS = {"start": "left", "end": "right", "justify": "both"}

_JC_PART_RE = re.compile(
    r"^word/(document|styles|numbering|header\d*|footer\d*"
    r"|footnotes|endnotes|comments)\.xml$")
_JC_RE = re.compile(r'(<w:jc\b[^>]*?w:val=")([A-Za-z]+)(")')


def _normalize_docx_quirks(src_path, out_dir):
    """把 python-docx 不认识的枚举值改写为等价标准值，返回可用 docx 路径。

    根因（2026-09-28 实测）：LibreOffice 6.4 导出 docx 时会把左/右对齐写成
    `w:jc w:val="start"` / `"end"`（ECMA-376 第 4 版的取值），而 python-docx 1.1.x 的
    WD_PARAGRAPH_ALIGNMENT 只有 left/right/center/both/distribute，读取时抛
    `ValueError: WD_PARAGRAPH_ALIGNMENT has no XML mapping for 'start'`，
    导致 .doc/.wps 转换出来的文档无法汇编（旧版 WPS COM 转换不会这样写，故此前未暴露）。

    未发现需要改写的值时**原样返回 src_path**（零开销，且不改动用户源文件）。
    """
    try:
        zf = zipfile.ZipFile(src_path, "r")
    except Exception:
        return src_path
    changed = {}
    try:
        for name in zf.namelist():
            if not _JC_PART_RE.match(name):
                continue
            data = zf.read(name).decode("utf-8", "replace")

            def repl(m):
                new = _JC_ALIAS.get(m.group(2))
                if new and new in _JC_KNOWN:
                    return m.group(1) + new + m.group(3)
                return m.group(0)

            new_data = _JC_RE.sub(repl, data)
            if new_data != data:
                changed[name] = new_data
        if not changed:
            return src_path
        work = tempfile.mkdtemp(prefix="quirk_")
        out = os.path.join(work, "normalized.docx")
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zf.infolist():
                payload = changed.get(item.filename)
                if payload is None:
                    payload = zf.read(item.filename)
                else:
                    payload = payload.encode("utf-8")
                zo.writestr(item, payload)
    finally:
        zf.close()
    return out


def convert_to_docx(src_path, out_dir):
    """src_path(.doc/.docx/.wps) -> .docx 路径。

    .doc/.wps 由 WPS / LibreOffice 转换；.docx 若能被 python-docx 正常打开则原样返回，
    否则（宏启用 .docm、模板等）先尝试轻量规范化（剥离宏、保留全部文本排版），
    失败再交转换引擎重存并兜底规范化。

    无论走哪条路径，最终都会做一次「枚举值归一化」（见 _normalize_docx_quirks），
    把 LibreOffice / 新版 Office 写出的 python-docx 不认识的属性值改写为等价标准值。
    """
    ext = os.path.splitext(src_path)[1].lower()
    if ext in (".doc", ".wps"):
        return _normalize_docx_quirks(_run_convert(src_path, out_dir, "docx"), out_dir)
    if ext == ".docx":
        if not _docx_needs_normalize(src_path):
            return _normalize_docx_quirks(src_path, out_dir)
        # 需要规范化：优先轻量 zip 级（保留文本/排版，剥离宏），失败再交 WPS
        try:
            return _normalize_docx_quirks(
                _normalize_docx_package(src_path, out_dir), out_dir)
        except Exception as e1:
            try:
                out = _run_convert(src_path, out_dir, "docx")
                # WPS 可能仍保留 macroEnabled 标记，再兜底规范化一次
                return _normalize_docx_quirks(
                    _normalize_docx_package(out, out_dir), out_dir)
            except Exception as e2:
                raise RuntimeError("docx 规范化失败：" +
                                   str(e1)[:120] + " | " + str(e2)[:120])
    raise ValueError(f"不支持的源格式：{ext}")


def convert_to_pdf(src_path, out_dir):
    """src_path(.docx/.doc/.wps) -> PDF 路径（用于预览）。"""
    return _run_convert(src_path, out_dir, "pdf")


def wps_can_convert(timeout=20):
    """能力自检：用临时 docx 实际试转一次，返回 (ok, msg)。

    ok=True 表示存在可用的转换引擎（Windows 的 WPS，或 Linux 的 LibreOffice）
    且可完成转换；ok=False 时 msg 给出原因，便于界面直接提示用户。

    timeout：自检允许的最长秒数，会一路传到引擎的子进程超时上。
    （2026-09-29 修：此前该参数被声明却从未使用，引擎内部默认 180 秒，
    一旦 LibreOffice 变慢，启动自检最多可卡约 3 分钟。）
    """
    name, exe = detect_converter()
    if not exe:
        return False, "未检测到可用的转换引擎（Linux 请安装 LibreOffice；Windows 请安装 WPS）。"
    try:
        from docx import Document
    except Exception:
        # 无法构造测试文件则跳过（不阻断程序，仅跳过自检）
        return True, ""
    td = tempfile.mkdtemp(prefix="conv_check_")
    try:
        src = os.path.join(td, "t.docx")
        Document().save(src)
        out = _run_convert(src, td, "pdf", timeout=timeout)
        ok = os.path.isfile(out) and os.path.getsize(out) > 0
        return ok, "" if ok else "转换未生成有效 PDF"
    except Exception as e:
        return False, (name + " 转换测试失败：" + str(e))[:300]
    finally:
        shutil.rmtree(td, ignore_errors=True)
