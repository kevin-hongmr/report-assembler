#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把 LibreOffice 官方 Linux deb tar.gz 解出为便携目录（供国产系统离线包用）。

LibreOffice 官方只发布 .deb/.rpm/.msi，没有官方"便携版"。要得到免 root 的便携树，
需把 .deb 逐个解开合并到一个目录：
    app/libreoffice/opt/libreoffice<版本>/program/soffice

用法：
    python tools/extract_libreoffice_debs.py <LibreOffice_X_Linux_aarch64_deb.tar.gz> \
        --out app/libreoffice

依赖：本机需有能解 ar 格式的 tar（Windows 自带 C:\\Windows\\System32\\tar.exe 即 bsdtar，
支持 ar + gz/xz/zstd；Git Bash 的 GNU tar 不支持 ar）。找不到可用 tar 时回退纯 Python
（仅支持 gz/xz 压缩的 data.tar，zstd 会失败并明确报错）。
"""
import argparse
import os
import shutil
import subprocess
import sys
import tarfile
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def find_bsdtar():
    """找一个支持 ar 格式的 tar（Windows 自带的 bsdtar）。"""
    cand = []
    win = os.environ.get("SystemRoot", r"C:\Windows")
    cand.append(os.path.join(win, "System32", "tar.exe"))
    for name in ("bsdtar",):
        p = shutil.which(name)
        if p:
            cand.append(p)
    for c in cand:
        if c and os.path.isfile(c):
            return c
    return None


def extract_deb_bsdtar(tar_cmd, deb_path, out):
    """bsdtar 两步解：先解出 ar 成员（data.tar.*），再解 data 层到 out。"""
    tmp = tempfile.mkdtemp(prefix="deb_")
    try:
        subprocess.run([tar_cmd, "-xf", deb_path, "-C", tmp],
                       check=True, capture_output=True)
        data_tar = None
        for fn in os.listdir(tmp):
            if fn.startswith("data.tar"):
                data_tar = os.path.join(tmp, fn)
                break
        if not data_tar:
            return False
        os.makedirs(out, exist_ok=True)
        subprocess.run([tar_cmd, "-xf", data_tar, "-C", out],
                       check=True, capture_output=True)
        return True
    except subprocess.CalledProcessError:
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def extract_deb_python(deb_path, out):
    """纯 Python 解 .deb（仅支持 gz/xz 的 data.tar；zstd 不支持）。"""
    # ar 格式解析
    with open(deb_path, "rb") as f:
        magic = f.read(8)
        if magic != b"!<arch>\n":
            return False
        data_bytes = None
        while True:
            hdr = f.read(60)
            if len(hdr) < 60:
                break
            name = hdr[0:16].decode("ascii", "replace").strip().rstrip("/")
            size_s = hdr[48:58].decode("ascii", "replace").strip()
            try:
                size = int(size_s)
            except ValueError:
                break
            body = f.read(size)
            if size % 2 == 1:
                f.read(1)  # 2 字节对齐
            if name.startswith("data.tar"):
                data_bytes = body
                # 只要 data 层；后续成员可跳过，但为简单继续读完（文件不大）
        if data_bytes is None:
            return False
    fd, tmp_tar = tempfile.mkstemp(suffix=".tar" + _comp_suffix(deb_path))
    os.close(fd)
    try:
        with open(tmp_tar, "wb") as f:
            f.write(data_bytes)
        mode = "r:" + _tarfile_mode(tmp_tar)
        os.makedirs(out, exist_ok=True)
        with tarfile.open(tmp_tar, mode) as tf:
            tf.extractall(out)
        return True
    except Exception:
        return False
    finally:
        if os.path.exists(tmp_tar):
            os.remove(tmp_tar)


def _comp_suffix(_unused):
    return ""


def _tarfile_mode(tmp_tar):
    # 根据内容嗅探压缩类型（gz/xz/zstd）
    with open(tmp_tar, "rb") as f:
        head = f.read(8)
    if head[:2] == b"\x1f\x8b":
        return "gz"
    if head[:6] == b"\xfd7zXZ\x00":
        return "xz"
    if head[:4] == b"\x28\xb5\x2f\xfd":
        raise RuntimeError("data.tar 为 zstd 压缩，纯 Python 不支持，请用 bsdtar")
    return ""  # 未压缩 tar


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deb_tar_gz")
    ap.add_argument("--out", default=os.path.join(ROOT, "app", "libreoffice"))
    opts = ap.parse_args()

    if not os.path.isfile(opts.deb_tar_gz):
        print("!! 找不到文件：%s" % opts.deb_tar_gz)
        return 1

    tar_cmd = find_bsdtar()
    use_bsdtar = tar_cmd is not None
    print("bsdtar：%s" % (tar_cmd or "未找到，将用纯 Python 回退"))

    out = opts.out
    if os.path.isdir(out):
        print("清空旧目录：%s" % out)
        shutil.rmtree(out)
    os.makedirs(out, exist_ok=True)

    tmp_root = tempfile.mkdtemp(prefix="lo_")
    failed = []
    try:
        print("解包外层 tar.gz ...")
        with tarfile.open(opts.deb_tar_gz, "r:gz") as tf:
            tf.extractall(tmp_root)
        debs_dir = None
        for base, _dirs, _files in os.walk(tmp_root):
            if os.path.basename(base) == "DEBS":
                debs_dir = base
                break
        if not debs_dir:
            print("!! 未找到 DEBS 目录")
            return 1
        debs = sorted(f for f in os.listdir(debs_dir) if f.endswith(".deb"))
        print("共 %d 个 .deb" % len(debs))
        for i, fn in enumerate(debs, 1):
            deb = os.path.join(debs_dir, fn)
            ok = False
            if use_bsdtar:
                ok = extract_deb_bsdtar(tar_cmd, deb, out)
            if not ok:
                ok = extract_deb_python(deb, out)
            if not ok:
                failed.append(fn)
            if i % 10 == 0 or i == len(debs):
                print("  进度 %d/%d" % (i, len(debs)))
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)

    if failed:
        print("以下 .deb 解包失败（%d 个）：" % len(failed))
        for fn in failed:
            print("  - %s" % fn)

    soffice = None
    for base, _dirs, files in os.walk(out):
        if "soffice" in files:
            soffice = os.path.join(base, "soffice")
            break
    print("=" * 60)
    print("soffice：%s" % (soffice or "未找到"))
    return 0 if soffice else 1


if __name__ == "__main__":
    sys.exit(main())
