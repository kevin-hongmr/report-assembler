#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""让「发行版为 Qt 6.4 编译的 fcitx Qt6 输入法插件」能在 PyQt6 自带的 Qt 6.7.3 上加载。

## 为什么需要它

fcitx 的 Qt6 输入法模块（发行版包名 `fcitx-frontend-qt6`，源码包 fcitx-qt5）是 Qt6
应用输入中文的**唯一**桥梁：系统里的 `fcitx-frontend-qt5` 只服务 Qt5 程序，而 PyQt6
自带的 Qt6 插件目录（`PyQt6/Qt6/plugins/platforminputcontexts/`）只有 ibus / compose，
没有 fcitx；Qt 6 又已彻底移除 XIM 支持（`libQt6XcbQpa.so.6` 中无任何 `_XIM_*` 符号），
所以既不能靠 Qt5 插件、也不能靠 XIM 兜底。国产系统（银河麒麟 V10 SP1）上的 fcitx4
是搜狗输入法底座，短期内不会换成 ibus。

现成可用的 Qt6 fcitx 插件只有 Debian bookworm 版（`fcitx-frontend-qt6 1.2.7-2+b7`，
按 Qt 6.4.2 编译）。它在本机会连撞两堵墙，本工具用两处**最小字节改写**把两堵墙都拆掉。

## 障碍一：libstdc++ 版本需求 GLIBCXX_3.4.29

该插件用 GCC 12 编译，对 `libstdc++.so.6` 有一条版本需求 `GLIBCXX_3.4.29`；麒麟自带
libstdc++ 由 GCC 9.3 构建，最高只到 `GLIBCXX_3.4.28`，动态链接器直接拒绝加载：

    libstdc++.so.6: version `GLIBCXX_3.4.29' not found

两条死路已实测排除：
- **LD_PRELOAD 垫片无用**：glibc 的版本需求是"按文件"校验的，链接器只在名为
  `libstdc++.so.6` 的对象里查这个版本节点，不搜预载对象（`LD_DEBUG=versions` 已证）。
- **整包换新版 libstdc++ 不可行**：Debian bookworm 的 `libstdc++.so.6.0.30` 自身需求
  `GLIBC_2.32/2.33/2.34/2.36`，麒麟 glibc 2.31 跑不起来。

**改法**：插件实际只引用了该版本节点下的**一个**符号
`_ZSt28__throw_bad_array_new_lengthv`（`std::__throw_bad_array_new_length()`，GCC 11
引入）。把它就地改名为麒麟具备的 `_ZSt17__throw_bad_allocv`（`std::__throw_bad_alloc()`，
版本节点 `GLIBCXX_3.4`），再把 `.gnu.version_r` 里对应的那条版本需求改指 `GLIBCXX_3.4`。
两者都是 `[[noreturn]]` 的抛出辅助，`std::bad_array_new_length` 本就派生自
`std::bad_alloc`，语义差异仅在" `new[]` 长度计算溢出"这一实际不可达分支上。

（备选的语义精确写法——改名为 `__cxa_throw_bad_array_new_length`，版本 `CXXABI_1.3.8`
——不可行：该版本名字符串不在本插件的 `.dynstr` 里，而 `.dynstr` 无法在不重建节表的情况下
就近取用。）

## 障碍二：Qt 6.4 与 6.7 的私有 API 符号不匹配

插件按 Qt 6.4.2 编译，引入了 `QWindowSystemInterface::handleExtendedKeyEvent` 的 12 参
形式；Qt 6.7 已**删掉其尾部第 12 个参数**（`tryShortcutOverride`），新导出为 11 参形式
（同一版本节点 `Qt_6_PRIVATE_API`）：

    旧（插件导入）  ...KeyboardModifierEEjjjRK7QString b t b   ← 多一个尾部 bool
    新（Qt 6.7 导出）...KeyboardModifierEEjjjRK7QString b t

未命中即报：

    undefined symbol: _ZN22QWindowSystemInterface22handleExtendedKeyEvent...btb,
        version Qt_6_PRIVATE_API

**改法**：把插件 `.dynstr` 中该导入符号名**截短一个字节**（去掉尾部 `b`），指向 Qt 6.7
实际导出的 11 参符号。这是**严格 ABI 兼容**的：两者寄存器传参完全相同（前 8 个形参），
栈参数为 `QString const&`、`bool autorep`、`unsigned short count`，多出的旧参数
`tryShortcutOverride` 位于**栈参数最末位**，被调用方直接忽略，不影响前面任何实参的落位。
符号所属版本节点 `Qt_6_PRIVATE_API` 保持不变即可正确绑定。

`handleExtendedKeyEvent` 在插件里只被 `QFcitxPlatformInputContext::forwardEvent()` 调用
（把输入法消费掉的按键转回窗口），即输入法的常规工作路径，因此必须能解析。

## 安全性

两处改写都只落在 `.dynstr` 的字符串字节上，不改变任何节的大小与偏移。改写前会校验：
- 目标字符串是 `.dynstr` 中的**独立条目**（前后均为 NUL），且全节唯一；
- 它不是其它更长条目的**后缀**（避免共享后缀存储被误改）；
- **没有任何 `.dynsym` 条目的 `st_name` 落在该字符串区间内部**（避免踩到指向同一串
  中段偏移的符号）；
- 新名不短于 1 字节、不长于旧名（长于旧名需扩节，本工具拒绝）。

用法：
    python3 tools/patch_qt6_ime_plugin.py <上游插件.so> [输出.so]
    python3 tools/patch_qt6_ime_plugin.py --check <已处理的.so>
"""

import os
import struct
import sys

# ---- 修改 1：老 libstdc++ 缺失的符号 → 语义等价且老 libstdc++ 具备的替代符号 ----
SYM_FROM = b"_ZSt28__throw_bad_array_new_lengthv"
SYM_TO = b"_ZSt17__throw_bad_allocv"
VER_FROM = b"GLIBCXX_3.4.29"      # 上游要求、老 libstdc++ 没有
VER_TO = b"GLIBCXX_3.4"           # 老 libstdc++ 具备

# ---- 修改 2：Qt 6.4 的 handleExtendedKeyEvent(12 参) → Qt 6.7 的 (11 参) ----
_QT_HKE_PREFIX = b"_ZN22QWindowSystemInterface22handleExtendedKeyEvent"
_QT_HKE_SIG = (b"EP7QWindowmN6QEvent4TypeEi6QFlagsIN2Qt16KeyboardModifierEE"
               b"jjjRK7QStringbt")
QT_SYM_FROM = _QT_HKE_PREFIX + _QT_HKE_SIG + b"b"
QT_SYM_TO = _QT_HKE_PREFIX + _QT_HKE_SIG
QT_SYM_VER = b"Qt_6_PRIVATE_API"   # 新旧符号同属该版本节点，无需改动

ELF_MAGIC = b"\x7fELF"


# --------------------------------------------------------------------------- ELF

def _sections(data):
    """解析 ELF64 节表，返回 [(名称, sh_offset, sh_size, sh_type, sh_entsize)]。"""
    if data[:4] != ELF_MAGIC:
        raise ValueError("不是 ELF 文件")
    if data[4] != 2:
        raise ValueError("仅支持 64 位 ELF")
    e_shoff, = struct.unpack_from("<Q", data, 0x28)
    e_shentsize, e_shnum, e_shstrndx = struct.unpack_from("<HHH", data, 0x3A)
    if not e_shoff or not e_shnum:
        raise ValueError("没有节表")
    raw = []
    for i in range(e_shnum):
        off = e_shoff + i * e_shentsize
        sh_name, sh_type = struct.unpack_from("<II", data, off)
        sh_offset, sh_size = struct.unpack_from("<QQ", data, off + 0x18)
        sh_entsize, = struct.unpack_from("<Q", data, off + 0x38)
        raw.append((sh_name, sh_type, sh_offset, sh_size, sh_entsize))
    _, _, str_off, str_size, _ = raw[e_shstrndx]
    names = data[str_off:str_off + str_size]
    return [(names[sh_name:names.find(b"\x00", sh_name)].decode("latin-1"),
             sh_offset, sh_size, sh_type, sh_entsize)
            for sh_name, sh_type, sh_offset, sh_size, sh_entsize in raw]


def _find_section(secs, name):
    for n, off, size, _t, _e in secs:
        if n == name:
            return off, size
    raise ValueError("找不到节 %s" % name)


def _section_entsize(secs, name):
    for n, _off, _size, _t, e in secs:
        if n == name:
            return e
    raise ValueError("找不到节 %s" % name)


def _dynstr_entries(blob):
    """枚举 .dynstr 中每个字符串条目的 (相对偏移, 内容)。

    不能直接用 index()：字符串表允许把 "GLIBCXX_3.4" 作为 "GLIBCXX_3.4.18" 的前缀
    共享存储，必须按"独立条目"（前是 NUL、后也是 NUL）来找。
    """
    out = []
    i, n = 0, len(blob)
    while i < n:
        end = blob.find(b"\x00", i)
        if end < 0:
            break
        if end > i:
            out.append((i, blob[i:end]))
        i = end + 1
    return out


def _dynstr_entry_offset(data, secs, needle, what="字符串"):
    """返回 needle 作为**独立字符串条目**在 .dynstr 内的相对偏移。"""
    off, size = _find_section(secs, ".dynstr")
    blob = data[off:off + size]
    hits = [rel for rel, s in _dynstr_entries(blob) if s == needle]
    if not hits:
        raise ValueError(".dynstr 中没有独立的%s条目 %r" % (what, needle))
    if len(hits) > 1:
        raise ValueError(".dynstr 中%s条目 %r 出现 %d 次（应为 1）" % (what, needle, len(hits)))
    return hits[0]


def _dynsym_name_offsets(data, secs):
    """全部 .dynsym 条目的 st_name 相对 .dynstr 的偏移集合。"""
    off, size = _find_section(secs, ".dynsym")
    entsize = _section_entsize(secs, ".dynsym") or 24
    return {struct.unpack_from("<I", data, p)[0] for p in range(off, off + size, entsize)}


def _check_rename_safe(data, secs, needle, old):
    """确认把 needle 原地改名为 old（更短或等长的同类名字）不会破坏其它字符串。"""
    if not old:
        raise ValueError("替代名不能为空")
    if len(old) > len(needle):
        raise ValueError("替代名 %r 比原名长，需扩节，本工具不支持" % old)
    off, size = _find_section(secs, ".dynstr")
    blob = data[off:off + size]
    cnt = blob.count(needle)
    if cnt != 1:
        raise ValueError(".dynstr 中 %r 出现 %d 次（应为 1）" % (needle, cnt))
    rel = blob.index(needle)
    if blob[rel + len(needle)] != 0x00:
        raise ValueError("%r 不是独立字符串（其后还有字节）" % needle)
    for _r, s in _dynstr_entries(blob):
        if s.endswith(needle) and s != needle:
            raise ValueError("%r 是更长字符串 %r 的后缀，不能原地改名" % (needle, s))
    # 没有任何符号的名字从该串的"中段"开始
    inner = {o for o in _dynsym_name_offsets(data, secs)
             if rel < o < rel + len(needle)}
    if inner:
        raise ValueError("%r 区间内有其它符号名起点偏移 %s，不能原地改名"
                         % (needle, sorted(inner)))
    return rel


def elf_hash(name):
    """glibc 的 ELF 版本名哈希算法。"""
    h = 0
    for c in name:
        h = ((h << 4) + c) & 0xFFFFFFFF
        g = h & 0xF0000000
        if g:
            h ^= g >> 24
        h &= ~g
    return h & 0xFFFFFFFF


def _iter_vernaux(data, secs):
    """遍历 .gnu.version_r，产出 (Vernaux 的文件偏移, vn_file 名, vna_hash,
    vna_other, vna_name 偏移, 该条目的版本名)。

    偏移语义：Verneed.vn_next / Vernaux.vna_next 都是**相对当前条目起点**的偏移，
    为 0 表示链尾，必须 break，不能回绕到 0 继续循环。
    """
    doff, _dsize = _find_section(secs, ".dynstr")
    off, size = _find_section(secs, ".gnu.version_r")
    blob = data[off:off + size]

    def s_at(rel):
        s = doff + rel
        return data[s:data.find(b"\x00", s)].decode("latin-1")

    pos = 0
    while pos < size:
        _vn_version, vn_cnt, vn_file, vn_aux, vn_next = struct.unpack_from("<HHIII", blob, pos)
        apos = pos + vn_aux
        for _ in range(vn_cnt):
            vna_hash, _vna_flags, vna_other, vna_name, vna_next = struct.unpack_from(
                "<IHHII", blob, apos)
            yield (off + apos, s_at(vn_file), vna_hash, vna_other, vna_name, s_at(vna_name))
            if not vna_next:
                break
            apos += vna_next
        if not vn_next:
            break
        pos += vn_next


def _version_of(data, secs):
    """把 .gnu.version 的版本索引表与 .gnu.version_r 的名字对应起来，
    返回 {dynsym 索引: 版本名}。"""
    voff, vsize = _find_section(secs, ".gnu.version")
    idx2name = {aux[3]: aux[-1] for aux in _iter_vernaux(data, secs)}
    out = {}
    for i in range(vsize // 2):
        idx, = struct.unpack_from("<H", data, voff + i * 2)
        if idx & 0x8000:      # VER_NDX_LOCAL / 隐藏位
            continue
        out[i] = idx2name.get(idx & 0x7FFF)
    return out


def dynsym_index_of(data, secs, symname):
    """按名字找动态符号表下标（未找到返回 None）。"""
    doff, _ = _find_section(secs, ".dynstr")
    off, size = _find_section(secs, ".dynsym")
    entsize = _section_entsize(secs, ".dynsym") or 24
    for i, p in enumerate(range(off, off + size, entsize)):
        st_name, = struct.unpack_from("<I", data, p)
        s = doff + st_name
        if data[s:data.find(b"\x00", s)] == symname:
            return i
    return None


# --------------------------------------------------------------------------- 主流程

def _rename_in_dynstr(data, secs, old_name, new_name):
    rel = _check_rename_safe(data, secs, old_name, new_name)
    dynstr_off, _ = _find_section(secs, ".dynstr")
    pad = len(old_name) - len(new_name)
    data[dynstr_off + rel:dynstr_off + rel + len(old_name)] = new_name + b"\x00" * pad
    return rel


def patch(src_path, dst_path):
    """两处改写：libstdc++ 符号/版本需求 + Qt 私有符号名截短。"""
    with open(src_path, "rb") as f:
        data = bytearray(f.read())
    secs = _sections(data)
    _find_section(secs, ".gnu.version_r")   # 存在性检查

    # ---- 修改 1a：符号改名 ----
    rel_from = _check_rename_safe(data, secs, SYM_FROM, SYM_TO)
    dynstr_off, _ = _find_section(secs, ".dynstr")
    data[dynstr_off + rel_from:dynstr_off + rel_from + len(SYM_FROM)] = \
        SYM_TO + b"\x00" * (len(SYM_FROM) - len(SYM_TO))

    # ---- 修改 1b：版本需求 GLIBCXX_3.4.29 → GLIBCXX_3.4 ----
    rel_ver_to = _dynstr_entry_offset(data, secs, VER_TO, "版本名")
    auxes = list(_iter_vernaux(data, secs))
    hash_to = None
    for _off_a, _f, vna_hash, _o, _n, name in auxes:
        if name == VER_TO.decode():
            hash_to = vna_hash
    if hash_to is None:
        hash_to = elf_hash(VER_TO)

    changed = 0
    for off_a, _f, _h, _o, _n, name in auxes:
        if name != VER_FROM.decode():
            continue
        struct.pack_into("<I", data, off_a, hash_to)            # vna_hash
        struct.pack_into("<I", data, off_a + 8, rel_ver_to)     # vna_name
        changed += 1
    if changed != 1:
        raise ValueError("版本需求 %r 命中 %d 条（应为 1）" % (VER_FROM.decode(), changed))

    # ---- 修改 2：Qt 私有符号名截短一个字节（12 参 → 11 参） ----
    _rename_in_dynstr(data, secs, QT_SYM_FROM, QT_SYM_TO)

    with open(dst_path, "wb") as f:
        f.write(data)
    return {"symbol": SYM_FROM.decode(), "symbol_new": SYM_TO.decode(),
            "version": VER_FROM.decode(), "version_new": VER_TO.decode(),
            "qt_symbol_from": QT_SYM_FROM.decode(), "qt_symbol_to": QT_SYM_TO.decode()}


def version_needs(path):
    """该 ELF 实际声明的全部版本需求名（按结构解析，不做全文搜索）。"""
    with open(path, "rb") as f:
        data = f.read()
    return {aux[-1] for aux in _iter_vernaux(data, _sections(data))}


def check(path):
    """核对处理结果。返回问题列表（空列表表示通过）。

    注意：不能用"全文搜索字符串"判断版本需求——.dynstr 里没人引用的旧名字串会保留，
    属无害残留，必须按 .gnu.version_r 结构解析才准。
    """
    with open(path, "rb") as f:
        data = f.read()
    secs = _sections(data)
    problems = []

    needs = {aux[-1] for aux in _iter_vernaux(data, secs)}
    if VER_FROM.decode() in needs:
        problems.append("版本需求中仍存在 %s" % VER_FROM.decode())
    if VER_TO.decode() not in needs:
        problems.append("版本需求中缺少 %s" % VER_TO.decode())
    if SYM_FROM in data:
        problems.append("仍存在旧符号名 %s" % SYM_FROM.decode())
    if SYM_TO not in data:
        problems.append("未找到替代符号 %s" % SYM_TO.decode())

    # Qt 符号：旧名（12 参）应已消失，新名（11 参）存在，且版本节点不变
    if QT_SYM_FROM in data:
        problems.append("Qt 符号仍为 12 参旧名（截短未生效）")
    if QT_SYM_TO not in data:
        problems.append("未找到 Qt 11 参符号名")
    i = dynsym_index_of(data, secs, QT_SYM_TO)
    if i is None:
        problems.append("Qt 11 参符号不在动态符号表中")
    else:
        ver = _version_of(data, secs).get(i)
        if ver != QT_SYM_VER.decode():
            problems.append("Qt 符号版本节点为 %r，期望 %r" % (ver, QT_SYM_VER.decode()))
    return problems


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 2
    if args[0] == "--check":
        if len(args) < 2:
            print("用法：--check <文件>")
            return 2
        problems = check(args[1])
        for p in problems:
            print("  ✗", p)
        print("RESULT:", "OK" if not problems else "有问题")
        return 1 if problems else 0
    src = args[0]
    dst = args[1] if len(args) > 1 else os.path.splitext(src)[0] + "_kylin.so"
    info = patch(src, dst)
    print("已生成:", dst, "(%d 字节)" % os.path.getsize(dst))
    print("  [1] 符号  %s → %s" % (info["symbol"], info["symbol_new"]))
    print("      版本  %s → %s" % (info["version"], info["version_new"]))
    print("  [2] Qt 符号 %s" % info["qt_symbol_from"])
    print("            → %s" % info["qt_symbol_to"])
    problems = check(dst)
    for p in problems:
        print("  ✗", p)
    print("RESULT:", "OK" if not problems else "有问题")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
