#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从麒麟官方源取 LibreOffice 并解为便携树（免 root，用于国产系统离线包）。

背景（2026-09-28）：TDF 官方 aarch64 版 LibreOffice 25.8.7 经 ELF VERNEED 实测
最高要求 GLIBC 2.34，而银河麒麟 V10 SP1 只有 glibc 2.31 → 无法运行（oosplash 报
"version `GLIBC_2.33' not found"）。麒麟官方源自带 LibreOffice 1:6.4.2-0kylin3，
libreoffice-core 声明 libc6 (>= 2.29)，与本机兼容，故改用它作为便携转换引擎来源。

本工具做三件事：
  1) 拉取并解析麒麟源的 Packages 索引（可给多个 suite）；
  2) 解析出「种子包 + 其 Depends 闭包」（可限深度、可跳过基础系统包/图标主题包）；
  3) 下载所需 .deb（不装、不改系统），并用 dpkg-deb -x 解到指定目录（便携树根）。

用法示例：
  python tools/fetch_kylin_libreoffice.py --arch aarch64 \
      --out build/lo_kylin64/debs --extract-dir build/lo_kylin64/lo \
      --depth 1

依赖：本机 dpkg-deb（解 .deb）与网络（curl 或 urllib）。
"""
import argparse
import gzip
import os
import re
import shutil
import ssl
import subprocess
import sys
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DEFAULT_SUITES = [
    "https://archive.kylinos.cn/kylin/KYLIN-ALL/dists/10.1/main/binary-arm64/Packages.gz",
    "https://archive.kylinos.cn/kylin/KYLIN-ALL/dists/10.1/restricted/binary-arm64/Packages.gz",
    "https://archive.kylinos.cn/kylin/KYLIN-ALL/dists/10.1/universe/binary-arm64/Packages.gz",
]
BASE_URL = "https://archive.kylinos.cn/kylin/KYLIN-ALL/"

# 转换引擎最小种子集（无头 Writer + 核心 + 公共文件 + UNO 运行时）
DEFAULT_SEEDS = [
    "libreoffice-writer-nogui",
    "libreoffice-core-nogui",
    "libreoffice-common",
    "libreoffice-base-core",
    "ure",
    "uno-libs-private",
    "fonts-opensymbol",
]

# 不下也不解：基础系统包（目标机必有，随包会破坏 ABI/体积）
SKIP_PREFIX = (
    "libc6", "libc-bin", "libc-dev", "libgcc", "libstdc++", "linux-libc-dev",
    "base-files", "base-passwd", "dpkg", "debconf", "perl", "perl-base",
    "sensible-utils", "zlib1g", "libselinux", "libattr1", "libcap2", "libpcre",
    "gcc-", "libcrypt", "tzdata", "ucf", "init-system-helpers", "sysvinit",
    "adduser", "hostname", "mawk", "grep", "sed", "gawk", "apt", "login",
    "passwd", "mount", "util-linux", "bash", "coreutils", "libbz2", "liblzma",
    "libzstd", "libffi", "libuuid", "libblkid", "libmount", "libtinfo", "libncurses",
)
# 不下也不解：GUI 图标主题（无头转换用不到，体积大）
SKIP_EXACT = {"libreoffice-style-colibre", "libreoffice-style-tango",
              "libreoffice-style-breeze", "libreoffice-style-elementary",
              "libreoffice-style-galaxy", "libreoffice-style-hicontrast",
              "libreoffice-style-sifr", "libreoffice-style-oxygen",
              # GUI 版核心（与 -nogui 互斥，无头转换只需 -nogui 版）
              "libreoffice-core"}


def fetch(url, dest=None, quiet=False):
    """下载 url；dest 为 None 时返回 bytes。

    优先用 curl：内网/代理环境常见自签证书，urllib 的严格校验会直接失败
    （麒麟源在本机即报 CERTIFICATE_VERIFY_FAILED），curl 走系统 CA 更稳。
    """
    data = None
    if shutil.which("curl"):
        cmd = ["curl", "-sSL", "--fail", "--connect-timeout", "30", url]
        if dest:
            cmd += ["-o", dest]
        r = subprocess.run(cmd, capture_output=True)
        if r.returncode == 0:
            if dest:
                if not quiet:
                    print("  下载 %-52s %6.1f MB"
                          % (os.path.basename(dest), os.path.getsize(dest) / 1e6))
                return dest
            return r.stdout
    if dest is None:
        ctx = ssl._create_unverified_context()
        req = urllib.request.Request(url, headers={"User-Agent": "curl/7.68"})
        with urllib.request.urlopen(req, timeout=180, context=ctx) as r:
            data = r.read()
    else:
        ctx = ssl._create_unverified_context()
        req = urllib.request.Request(url, headers={"User-Agent": "curl/7.68"})
        with urllib.request.urlopen(req, timeout=180, context=ctx) as r:
            data = r.read()
        with open(dest, "wb") as f:
            f.write(data)
        if not quiet:
            print("  下载 %-52s %6.1f MB"
                  % (os.path.basename(dest), len(data) / 1e6))
        return dest
    return data


def parse_index(raw):
    """解析 Packages 文本 → {包名: {version, filename, size, depends[]}}。"""
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    text = raw.decode("utf-8", "replace")
    out = {}
    for blk in text.split("\n\n"):
        m = re.search(r"^Package: (\S+)$", blk, re.M)
        if not m:
            continue
        name = m.group(1)
        ver = re.search(r"^Version: (\S+)$", blk, re.M)
        fn = re.search(r"^Filename: (\S+)$", blk, re.M)
        sz = re.search(r"^Size: (\d+)$", blk, re.M)
        dep = re.search(r"^Depends: (.+)$", blk, re.M)
        # 同包多版本时保留版本号最大者（索引内一般只有一个）
        if name in out:
            continue
        out[name] = {
            "version": ver.group(1) if ver else "",
            "filename": fn.group(1) if fn else "",
            "size": int(sz.group(1)) if sz else 0,
            "depends": dep.group(1) if dep else "",
        }
    return out


def dep_names(dep_field):
    """从 Depends 字段取出候选包名（处理 | 备选，取各组第一个）。"""
    names = []
    for group in dep_field.split(","):
        group = group.strip()
        if not group:
            continue
        first = group.split("|")[0].strip()
        m = re.match(r"([A-Za-z0-9][A-Za-z0-9+.\-]*)", first)
        if m:
            names.append(m.group(1))
    return names


def skipped(name):
    if name in SKIP_EXACT:
        return True
    low = name.lower()
    return any(low.startswith(p) for p in SKIP_PREFIX)


def resolve_closure(info, seeds, depth):
    """种子 + Depends 闭包（按深度限制；跳过基础系统包与图标主题）。"""
    wanted = {}          # name -> reason
    queue = [(s, 0, "seed") for s in seeds]
    missing = []
    while queue:
        name, d, reason = queue.pop(0)
        if name in wanted:
            continue
        if name not in info:
            if reason == "seed":
                missing.append(name)
            continue
        if skipped(name):
            continue
        wanted[name] = reason
        if d >= depth:
            continue
        for dn in dep_names(info[name]["depends"]):
            if dn not in wanted and not skipped(dn):
                queue.append((dn, d + 1, name))
    return wanted, missing


def download_debs(info, names, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    got, total = [], 0
    for name in sorted(names):
        meta = info[name]
        if not meta["filename"]:
            continue
        dest = os.path.join(out_dir, os.path.basename(meta["filename"]))
        if os.path.isfile(dest) and os.path.getsize(dest) == meta["size"]:
            got.append(dest)
            continue
        url = BASE_URL + meta["filename"]
        try:
            fetch(url, dest)
        except Exception as e:
            print("  !! 下载失败 %s: %s" % (name, str(e)[:120]))
            continue
        got.append(dest)
        total += meta["size"]
    return got


def extract_debs(debs, out_dir):
    if not shutil.which("dpkg-deb"):
        raise RuntimeError("未找到 dpkg-deb，无法解 .deb")
    os.makedirs(out_dir, exist_ok=True)
    for deb in debs:
        r = subprocess.run(["dpkg-deb", "-x", deb, out_dir],
                           capture_output=True, text=True)
        if r.returncode != 0:
            print("  !! 解包失败 %s: %s" % (os.path.basename(deb),
                                          (r.stderr or "")[:160]))
            return False
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--arch", default="aarch64", help="仅用于提示，实际由 --suite 决定")
    ap.add_argument("--suite", action="append", default=[],
                    help="Packages.gz 地址，可重复；缺省用麒麟 10.1 main/restricted/universe")
    ap.add_argument("--seeds", default=",".join(DEFAULT_SEEDS))
    ap.add_argument("--extra", default="",
                    help="额外追加的包（逗号分隔）；用于补齐 ldd 报出的传递依赖，"
                         "如 libboost-filesystem1.71.0,libraptor2-0,librasqal3")
    ap.add_argument("--depth", type=int, default=1, help="Depends 展开深度（0=只要种子）")
    ap.add_argument("--out", default=os.path.join(ROOT, "build", "lo_kylin64", "debs"))
    ap.add_argument("--extract-dir", default=None, help="把下载的 deb 解到该目录（便携树根）")
    ap.add_argument("--index-only", action="store_true", help="只解析索引并打印，不下载")
    ap.add_argument("--list", action="store_true", help="打印闭包清单后退出")
    opts = ap.parse_args()

    suites = opts.suite or DEFAULT_SUITES
    info = {}
    for url in suites:
        try:
            raw = fetch(url)
        except Exception as e:
            print("跳过索引 %s（%s）" % (url, str(e)[:80]))
            continue
        part = parse_index(raw)
        for k, v in part.items():
            info.setdefault(k, v)
        print("索引 %-70s %d 个包" % (url.split("dists/")[-1], len(part)))
    if not info:
        print("!! 未取到任何索引")
        return 1

    seeds = [s.strip() for s in opts.seeds.split(",") if s.strip()]
    seeds += [s.strip() for s in opts.extra.split(",") if s.strip()]
    wanted, missing = resolve_closure(info, seeds, opts.depth)
    if missing:
        print("!! 索引中找不到种子包：%s" % ", ".join(missing))
    total = sum(info[n]["size"] for n in wanted)
    print("\n需要 %d 个包，合计 %.1f MB：" % (len(wanted), total / 1e6))
    for n in sorted(wanted):
        print("   %-32s %-22s %6.1fMB" % (n, info[n]["version"], info[n]["size"] / 1e6))
    if opts.list or opts.index_only:
        return 0

    print("\n开始下载 ...")
    debs = download_debs(info, wanted, opts.out)
    print("已下载/已存在 %d 个 .deb -> %s" % (len(debs), opts.out))
    if opts.extract_dir:
        print("解包到 %s ..." % opts.extract_dir)
        if not extract_debs(debs, opts.extract_dir):
            return 1
        print("解包完成。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
