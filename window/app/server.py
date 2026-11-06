# -*- coding: utf-8 -*-
"""服务器版：在 Windows 上启动 HTTP 服务，局域网内其他电脑用浏览器访问。

- 汇编逻辑复用单机版（``app.core.pipeline`` / ``assembler`` / WPS COM 转换）；
- 同一时刻只处理一次汇编（WPS COM 不适合并发），其余请求排队；
- 结果 docx 以临时文件形式保留，供浏览器下载。

启动：双击「启动服务器版.bat」，或 ``python server.py --port 12362``。
"""
import argparse
import base64
import json
import os
import socket
import sys
import tempfile
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote

# 以 `python app/server.py` 直接运行时，需把项目根加入 sys.path
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from app.core import assembler, converter, pipeline

HERE = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(HERE, "web")
WORK_ROOT = os.path.join(tempfile.gettempdir(), "doc_assembler_server")

_jobs = {}                       # token -> {path, name, ts}
_jobs_lock = threading.Lock()
_asm_lock = threading.Lock()     # 串行化汇编（WPS COM）
RESULT_TTL = 3600                # 结果保留 1 小时


def _lan_ip():
    """取本机在局域网中的 IP（不实际发包）。"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def _cleanup_jobs():
    now = time.time()
    with _jobs_lock:
        for k in [k for k, v in _jobs.items() if now - v["ts"] > RESULT_TTL]:
            p = _jobs.pop(k)["path"]
            try:
                os.remove(p)
            except Exception:
                pass


def _derive_unit(name):
    """从**原始文件名**推断单位（「单位-标题」取前段；否则用整个文件名）。

    注意：不能用服务器临时文件名（src01.docx）推断，否则单位会变成 src01。
    """
    base = os.path.splitext(os.path.basename(name or ""))[0].strip()
    if "-" in base:
        return base.split("-", 1)[0].strip()
    return base


def _derive_title(name):
    """从原始文件名推断标题（「单位-标题」取后段；否则用整个文件名）。"""
    base = os.path.splitext(os.path.basename(name or ""))[0].strip()
    if "-" in base:
        return base.split("-", 1)[1].strip()
    return base


class Handler(BaseHTTPRequestHandler):
    server_version = "DocAssemblerServer/1.0"

    def log_message(self, fmt, *args):
        print("[%s] %s" % (time.strftime("%H:%M:%S"), fmt % args))

    # ---------------- 基础响应 ----------------
    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except Exception:
            pass

    def _json(self, code, obj):
        self._send(code, obj, "application/json; charset=utf-8")

    # ---------------- GET ----------------
    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            return self._file(os.path.join(WEB_DIR, "index.html"), "text/html; charset=utf-8")

        if path in ("/icon.png", "/favicon.ico"):
            return self._file(os.path.join(WEB_DIR, "icon.png"), "image/png")

        if path == "/api/health":
            try:
                eng = converter.detect_engine()
            except Exception as e:
                eng = "error: %s" % e
            return self._json(200, {"ok": True, "engine": eng,
                                    "host": socket.gethostname(), "ip": _lan_ip()})

        if path.startswith("/api/download/"):
            token = path.rsplit("/", 1)[-1]
            with _jobs_lock:
                job = _jobs.get(token)
            if not job or not os.path.isfile(job["path"]):
                return self._json(404, {"ok": False, "error": "文件不存在或已过期"})
            return self._file(job["path"],
                              "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                              download_name=job["name"])

        if path.startswith("/api/preview/"):
            parts = path.split("/")           # ['', 'api', 'preview', token, 'page_1.png']
            if len(parts) >= 5:
                token, fname = parts[3], os.path.basename(parts[4])
                with _jobs_lock:
                    job = _jobs.get(token)
                pdir = (job or {}).get("preview_dir")
                if pdir:
                    p = os.path.join(pdir, fname)
                    if os.path.isfile(p):
                        return self._file(p, "image/png")
            return self._json(404, {"ok": False, "error": "not found"})

        return self._json(404, {"ok": False, "error": "not found"})

    def _file(self, path, ctype, download_name=None):
        if not os.path.isfile(path):
            return self._json(404, {"ok": False, "error": "not found"})
        with open(path, "rb") as f:
            data = f.read()
        extra = {}
        if download_name:
            extra["Content-Disposition"] = "attachment; filename*=UTF-8''%s" % quote(download_name)
        self._send(200, data, ctype, extra)

    # ---------------- POST ----------------
    def do_POST(self):
        path = self.path.split("?")[0]
        if path not in ("/api/assemble", "/api/analyze"):
            return self._json(404, {"ok": False, "error": "not found"})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length) if length else b"{}"
            req = json.loads(raw.decode("utf-8"))
        except Exception as e:
            return self._json(400, {"ok": False, "error": "请求解析失败：%s" % e})

        try:
            if path == "/api/analyze":
                return self._json(200, self._analyze(req))
            return self._json(200, self._assemble(req))
        except Exception as e:
            return self._json(500, {"ok": False, "error": str(e)[:800]})

    def _save_files(self, files, job_dir):
        """把 base64 文件落盘，返回 [(显示名, 本地路径)]。"""
        out = []
        for i, f in enumerate(files, 1):
            name = os.path.basename(f.get("name") or ("file%d" % i))
            try:
                data = base64.b64decode(f.get("b64") or "")
            except Exception:
                data = b""
            if not data:
                continue
            ext = os.path.splitext(name)[1].lower() or ".docx"
            p = os.path.join(job_dir, "src%02d%s" % (i, ext))
            with open(p, "wb") as fh:
                fh.write(data)
            out.append((name, p))
        return out

    def _analyze(self, req):
        """识别每份文档的单位与标题（与单机版一致：读内容，文件名仅兜底）。"""
        files = req.get("files") or []
        if not files:
            raise RuntimeError("未收到文件")
        os.makedirs(WORK_ROOT, exist_ok=True)
        job_dir = os.path.join(WORK_ROOT, "an_" + uuid.uuid4().hex)
        os.makedirs(job_dir, exist_ok=True)
        items = self._save_files(files, job_dir)
        result = []
        with _asm_lock:                     # 可能触发 .doc/.wps 的 WPS COM 转换
            com_ready = False
            try:
                import pythoncom
                pythoncom.CoInitialize()
                com_ready = True
            except Exception:
                pass
            try:
                for name, p in items:
                    try:
                        dx = converter.convert_to_docx(p, job_dir)
                        info = assembler.analyze(dx)
                        unit = (info.get("unit") or "").strip() or _derive_unit(name)
                        title = (info.get("title") or "").strip() or _derive_title(name)
                        result.append({"name": name, "unit": unit, "title": title})
                    except Exception as e:
                        result.append({"name": name, "unit": _derive_unit(name),
                                       "title": _derive_title(name), "error": str(e)[:200]})
            finally:
                if com_ready:
                    try:
                        import pythoncom
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
        return {"ok": True, "items": result}

    def _assemble(self, req):
        files = req.get("files") or []
        if not files:
            raise RuntimeError("未收到文件")
        cfg = req.get("cfg") or {}
        options = req.get("options") or {}
        fmt = req.get("fmt") or None

        _cleanup_jobs()
        os.makedirs(WORK_ROOT, exist_ok=True)
        job_dir = os.path.join(WORK_ROOT, uuid.uuid4().hex)
        os.makedirs(job_dir, exist_ok=True)

        sources = []
        for i, f in enumerate(files, 1):
            name = os.path.basename(f.get("name") or ("file%d" % i))
            try:
                data = base64.b64decode(f.get("b64") or "")
            except Exception:
                data = b""
            if not data:
                continue
            ext = os.path.splitext(name)[1].lower() or ".docx"
            p = os.path.join(job_dir, "src%02d%s" % (i, ext))
            with open(p, "wb") as out:
                out.write(data)
            sources.append({"path": p,
                            "unit": (f.get("unit") or "").strip() or _derive_unit(name),
                            "title": (f.get("title") or "").strip() or _derive_title(name)})
        if not sources:
            raise RuntimeError("上传的文件为空")

        out_path = os.path.join(job_dir, "汇编结果.docx")
        preview_dir = os.path.join(job_dir, "preview")
        pages = 0

        # WPS COM 需串行；工作线程内须 CoInitialize
        with _asm_lock:
            com_ready = False
            try:
                import pythoncom
                pythoncom.CoInitialize()
                com_ready = True
            except Exception:
                pass
            try:
                pipeline.run_assemble(sources, out_path, cfg, options, job_dir, fmt=fmt)
                # 预览：docx -> PDF -> 页图（复用 WPS COM，同在本线程内）
                if req.get("preview", True):
                    try:
                        from app.core import preview as preview_mod
                        imgs = preview_mod.build_preview(out_path, preview_dir)
                        pages = len(imgs)
                    except Exception as e:
                        print("预览生成失败:", str(e)[:200])
            finally:
                if com_ready:
                    try:
                        import pythoncom
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        token = uuid.uuid4().hex
        with _jobs_lock:
            _jobs[token] = {"path": out_path, "name": "汇编结果.docx",
                            "preview_dir": preview_dir, "pages": pages, "ts": time.time()}
        return {"ok": True, "token": token, "name": "汇编结果.docx", "pages": pages}


def _tray_icon():
    """托盘图标：优先用 app/web/icon.png（本项目生成的服务器图标）。"""
    from PyQt6.QtGui import QIcon
    for p in (os.path.join(WEB_DIR, "icon.png"),
              os.path.join(os.path.dirname(HERE), "图标.png")):
        if os.path.isfile(p):
            return QIcon(p)
    return QIcon()


def _run_tray(port):
    """右下角系统托盘：双击打开界面，右键菜单可查看地址 / 退出。"""
    import webbrowser

    from PyQt6.QtWidgets import QApplication, QMenu, QSystemTrayIcon

    app = QApplication([])
    app.setQuitOnLastWindowClosed(False)

    ip = _lan_ip()
    url = "http://%s:%d" % (ip, port)
    tray = QSystemTrayIcon(_tray_icon())
    tray.setToolTip("汇报汇编程序 · 服务器版\n" + url)

    menu = QMenu()
    a_open = menu.addAction("打开界面（本机）")
    a_addr = menu.addAction("局域网地址：" + url)
    a_addr.setEnabled(False)
    menu.addSeparator()
    a_quit = menu.addAction("退出")
    tray.setContextMenu(menu)

    a_open.triggered.connect(lambda: webbrowser.open("http://127.0.0.1:%d" % port))
    a_quit.triggered.connect(app.quit)

    def _on_activated(reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            webbrowser.open("http://127.0.0.1:%d" % port)

    tray.activated.connect(_on_activated)
    tray.show()
    try:
        tray.showMessage("汇报汇编程序 · 服务器版",
                         "服务已启动，局域网访问：\n" + url,
                         QSystemTrayIcon.MessageIcon.Information, 5000)
    except Exception:
        pass

    print("托盘已启动；右键图标可查看地址或退出。")
    return app.exec()


def main():
    ap = argparse.ArgumentParser(description="汇报汇编程序 · 服务器版")
    ap.add_argument("--host", default="0.0.0.0", help="监听地址（默认 0.0.0.0，局域网可访问）")
    ap.add_argument("--port", type=int, default=12362, help="端口（默认 12362）")
    ap.add_argument("--console", action="store_true", help="控制台模式（不显示托盘图标）")
    opts = ap.parse_args()

    os.makedirs(WORK_ROOT, exist_ok=True)
    srv = ThreadingHTTPServer((opts.host, opts.port), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()

    ip = _lan_ip()
    print("=" * 56)
    print("  汇报汇编程序 · 服务器版  已启动")
    print("  本机访问：   http://127.0.0.1:%d" % opts.port)
    print("  局域网访问： http://%s:%d" % (ip, opts.port))
    print("  （同一局域网的电脑用浏览器打开上面第二个地址）")
    print("=" * 56)

    try:
        if not opts.console:
            try:
                return _run_tray(opts.port)
            except Exception as e:
                print("托盘启动失败，退回控制台模式：", e)
        while True:
            time.sleep(3600)
    except KeyboardInterrupt:
        print("\n已停止")
    finally:
        try:
            srv.shutdown()
        except Exception:
            pass
        srv.server_close()


if __name__ == "__main__":
    main()
