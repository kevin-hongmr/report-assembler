# -*- coding: utf-8 -*-
"""服务器版：在 Windows 上启动 HTTP 服务，局域网内其他电脑用浏览器访问。

- 汇编逻辑复用单机版（``app.core.pipeline`` / ``assembler`` / WPS COM 转换）；
- 同一时刻只处理一次汇编（WPS COM 不适合并发），其余请求排队；
- 结果 docx 以临时文件形式保留，供浏览器下载。

启动：双击「启动服务器.bat」，或 ``python server.py --port 8000``。
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

from app.core import pipeline, converter

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
        if path != "/api/assemble":
            return self._json(404, {"ok": False, "error": "not found"})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length) if length else b"{}"
            req = json.loads(raw.decode("utf-8"))
        except Exception as e:
            return self._json(400, {"ok": False, "error": "请求解析失败：%s" % e})

        try:
            return self._json(200, self._assemble(req))
        except Exception as e:
            return self._json(500, {"ok": False, "error": str(e)[:800]})

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
                            "unit": (f.get("unit") or "").strip(),
                            "title": (f.get("title") or "").strip()})
        if not sources:
            raise RuntimeError("上传的文件为空")

        out_path = os.path.join(job_dir, "汇编结果.docx")

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
            finally:
                if com_ready:
                    try:
                        import pythoncom
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        token = uuid.uuid4().hex
        with _jobs_lock:
            _jobs[token] = {"path": out_path, "name": "汇编结果.docx", "ts": time.time()}
        return {"ok": True, "token": token, "name": "汇编结果.docx"}


def main():
    ap = argparse.ArgumentParser(description="汇报汇编程序 · 服务器版")
    ap.add_argument("--host", default="0.0.0.0", help="监听地址（默认 0.0.0.0，局域网可访问）")
    ap.add_argument("--port", type=int, default=8000, help="端口（默认 8000）")
    opts = ap.parse_args()

    os.makedirs(WORK_ROOT, exist_ok=True)
    srv = ThreadingHTTPServer((opts.host, opts.port), Handler)
    ip = _lan_ip()
    print("=" * 56)
    print("  汇报汇编程序 · 服务器版  已启动")
    print("  本机访问：   http://127.0.0.1:%d" % opts.port)
    print("  局域网访问： http://%s:%d" % (ip, opts.port))
    print("  （同一局域网的电脑用浏览器打开上面第二个地址）")
    print("  按 Ctrl+C 停止服务")
    print("=" * 56)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    finally:
        srv.server_close()


if __name__ == "__main__":
    main()
