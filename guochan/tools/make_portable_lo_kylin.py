#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把「麒麟源 LibreOffice deb 解包目录」组装成自包含的便携树（免 root 转换引擎）。

为什么需要它：
  TDF 官方 aarch64 LibreOffice 25.8.7 实测最高要求 GLIBC 2.34，麒麟 V10 SP1（glibc 2.31）
  无法运行；改用麒麟官方源 1:6.4.2-0kylin3（libc6 >= 2.29）后，其第三方依赖（boost/clucene/
  orcus 等）在部分机器上缺失，需随包携带。

做法（关键点：LO 的二进制 RUNPATH 为 $ORIGIN）：
  1. 把 deb 解包目录里的 usr/lib/libreoffice/* 复制为 <out>/opt/libreoffice<ver>/；
  2. 用 ldd 迭代探测“在本机找不到”的共享库，从 deb 解包目录里取出（按 soname 命名、解引用软链），
     直接放进 program/ —— 因 RUNPATH=$ORIGIN，无需任何 LD_LIBRARY_PATH 或包装脚本；
  3. 只补“本机缺失”的库：已装于系统的库不随包，避免包体膨胀与 ABI 冲突；
  4. 剪掉 gallery 等与无头转换无关的资源。

用法：
  python tools/make_portable_lo_kylin.py \
      --stage build/lo_kylin64/stage --out build/lo_kylin64/lo
"""
import argparse
import os
import re
import shutil
import subprocess
import sys

SONAME_RE = re.compile(r"^\s+(\S+)\s+=>\s+not found")


def ldd_missing(paths, extra_paths):
    """返回这些 ELF 在本机解析不到的 soname 集合。"""
    env = dict(os.environ)
    if extra_paths:
        env["LD_LIBRARY_PATH"] = ":".join(extra_paths)
    missing = set()
    for p in paths:
        r = subprocess.run(["ldd", p], capture_output=True, text=True, env=env)
        for line in (r.stdout or "").splitlines():
            m = SONAME_RE.match(line)
            if m:
                missing.add(m.group(1))
    return missing


def elf_files(d):
    out = []
    for base, _dirs, files in os.walk(d):
        for fn in files:
            if fn.endswith((".so", ".bin")) or ".so." in fn or fn == "oosplash":
                out.append(os.path.join(base, fn))
    return out


def find_provider(stage, soname):
    """在 deb 解包目录里找提供该 soname 的文件。"""
    for d in ("usr/lib/aarch64-linux-gnu", "lib/aarch64-linux-gnu",
              "usr/lib", "lib"):
        root = os.path.join(stage, d)
        if not os.path.isdir(root):
            continue
        cand = os.path.join(root, soname)
        if os.path.exists(cand):
            return cand
    # 兜底：全树按文件名找
    for base, _dirs, files in os.walk(stage):
        if soname in files:
            return os.path.join(base, soname)
    return None


def move_aside(path, out_root):
    """把不需要的目录挪走。

    注意：本环境（以及部分安全策略环境）下 shutil.rmtree 会被 safe-delete 拦截并抛错，
    故一律用 mv 挪到 <out>/.pruned/ 下，不做真删除。
    """
    if not os.path.exists(path):
        return
    dest_dir = os.path.join(out_root, ".pruned")
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, os.path.basename(path.rstrip("/")))
    i = 1
    while os.path.exists(dest):
        dest = os.path.join(dest_dir, "%s.%d" % (os.path.basename(path.rstrip("/")), i))
        i += 1
    subprocess.run(["mv", path, dest], check=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--stage", required=True, help="deb 解包目录（含 usr/lib/libreoffice）")
    ap.add_argument("--out", required=True, help="便携树输出根（其下建 opt/libreoffice<ver>）")
    ap.add_argument("--name", default="libreoffice6.4", help="opt/ 下的目录名")
    ap.add_argument("--keep-gallery", action="store_true")
    opts = ap.parse_args()

    src = os.path.join(opts.stage, "usr", "lib", "libreoffice")
    if not os.path.isdir(src):
        print("!! 未找到 %s" % src)
        return 1
    dst = os.path.join(opts.out, "opt", opts.name)
    move_aside(dst, opts.out)
    os.makedirs(dst, exist_ok=True)

    print("复制 LibreOffice 主树 ...")
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks=True)
        else:
            shutil.copy2(s, d)

    # 剪掉与无头转换无关的大件
    for rel in ([] if opts.keep_gallery else ["share/gallery"]):
        p = os.path.join(dst, rel)
        if os.path.isdir(p):
            move_aside(p, opts.out)
            print("剪除 %s" % rel)

    # 关键：发行版打包的 fundamentalrc 里 BRAND_BASE_DIR 是**绝对路径**
    # （麒麟/Debian 均为 file:///usr/lib/libreoffice），搬走后会导致启动即抛
    # com::sun::star::uno::DeploymentException。TDF 官方可搬版本此处写的是
    # ${ORIGIN}/..，照此改回相对形式即可。
    fundamental = os.path.join(dst, "program", "fundamentalrc")
    if os.path.isfile(fundamental):
        with open(fundamental, "r", encoding="utf-8", errors="replace") as f:
            txt = f.read()
        new = re.sub(r"(?m)^BRAND_BASE_DIR=.*$", "BRAND_BASE_DIR=${ORIGIN}/..", txt)
        if new != txt:
            with open(fundamental, "w", encoding="utf-8") as f:
                f.write(new)
            print("已修正 program/fundamentalrc：BRAND_BASE_DIR=${ORIGIN}/..")
        # 顺带体检：报告树内其它残留的绝对安装路径（不影响启动者仅提示）
        leftovers = []
        for base, _dirs, files in os.walk(dst):
            for fn in files:
                if not fn.endswith((".rc", ".ini", ".conf")):
                    continue
                p = os.path.join(base, fn)
                try:
                    with open(p, "r", encoding="utf-8", errors="replace") as f:
                        if "/usr/lib/libreoffice" in f.read():
                            leftovers.append(os.path.relpath(p, dst))
                except Exception:
                    pass
        if leftovers:
            print("提示：以下配置仍含 /usr/lib/libreoffice 绝对路径（一般无害）：%s"
                  % ", ".join(sorted(leftovers)[:6]))

    prog = os.path.join(dst, "program")
    stage_libdirs = [os.path.join(opts.stage, "usr", "lib", "aarch64-linux-gnu"),
                     os.path.join(opts.stage, "lib", "aarch64-linux-gnu")]

    # 迭代补库：直到 ldd 不再报 not found
    # 注意：探测时只能把 program/ 放进搜索路径，绝不能把 stage 的库目录也放进去，
    # 否则 stage 里的库会被判定为“本机已有”，导致该补的库一个都补不上（曾踩此坑）。
    copied = set()
    for rnd in range(1, 6):
        missing = ldd_missing(elf_files(prog), [prog])
        todo = sorted(m for m in missing if m not in copied)
        if not todo:
            print("第 %d 轮：无缺失库，收敛。" % rnd)
            break
        print("第 %d 轮：补 %d 个库" % (rnd, len(todo)))
        for soname in todo:
            prov = find_provider(opts.stage, soname)
            if not prov:
                print("   !! 提供者不在 stage 内，需另行下载：%s" % soname)
                continue
            shutil.copyfile(prov, os.path.join(prog, soname))  # copyfile 解引用软链
            copied.add(soname)
            print("   + %-34s <- %s" % (soname, os.path.relpath(prov, opts.stage)))
    else:
        print("!! 5 轮后仍有缺失：%s" % ", ".join(sorted(
            ldd_missing(elf_files(prog), [prog]))))

    # 权限：可读、可执行（跨机复制常丢可执行位）
    for base, dirs, files in os.walk(opts.out):
        os.chmod(base, 0o755)
        for fn in files:
            p = os.path.join(base, fn)
            if os.path.islink(p):
                continue
            mode = os.stat(p).st_mode
            os.chmod(p, mode | 0o444 | (0o111 if mode & 0o100 else 0))

    # 包装脚本：把 program/ 加入 LD_LIBRARY_PATH。
    # 为什么必须这么做：LO 的二进制 RUNPATH 是 $ORIGIN，只对“直接依赖”生效；
    # 我们随包补进来的第三方库（如 libclucene-core）自己去找它的依赖
    # （libclucene-shared）时，不会继承调用者的 RUNPATH，于是报
    # "cannot open shared object file"。用包装脚本设 LD_LIBRARY_PATH 即解决。
    soffice = os.path.join(prog, "soffice")
    real = os.path.join(prog, "soffice.real")
    if os.path.isfile(soffice) and not os.path.isfile(real):
        shutil.move(soffice, real)
        with open(soffice, "w") as f:
            f.write('#!/bin/sh\n'
                    '# 便携 LibreOffice 启动包装（由 tools/make_portable_lo_kylin.py 生成）\n'
                    '# 作用：把本目录加入 LD_LIBRARY_PATH，使随包携带的第三方库及其\n'
                    '#       相互依赖都能被动态链接器找到（RUNPATH=$ORIGIN 不具传递性）。\n'
                    'HERE=$(cd "$(dirname "$0")" && pwd)\n'
                    'LD_LIBRARY_PATH="$HERE${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"\n'
                    'export LD_LIBRARY_PATH\n'
                    'exec "$HERE/soffice.real" "$@"\n')
        os.chmod(soffice, 0o755)
        os.chmod(real, 0o755)
        print("已生成包装脚本：program/soffice（原脚本保留为 soffice.real）")

    print("\n便携树：%s" % dst)
    print("soffice：%s（存在=%s 可执行=%s）"
          % (soffice, os.path.isfile(soffice), os.access(soffice, os.X_OK)))
    print("随包补齐的第三方库：%d 个" % len(copied))
    size = subprocess.run(["du", "-sh", dst], capture_output=True, text=True)
    print("体积：%s" % (size.stdout.strip() or "?"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
