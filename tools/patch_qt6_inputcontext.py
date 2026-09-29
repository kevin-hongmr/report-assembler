#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""让「按 Qt 6.4 编译的 Qt6 输入法插件」在 PyQt6 自带的 Qt 6.7.3 上真正可用。

## 问题（已定位到单条指令）

Qt 6.7 的 `QPlatformInputContext` 类在末尾新增了数据成员
`Qt::LayoutDirection m_inputDirection;`（偏移 16；Qt 6.4 无此成员，类大小为 16）。
同时它的构造函数会排入一个 **QueuedConnection 延迟调用**，在事件循环第一次迭代时执行：

    m_inputDirection = QLocale::system().textDirection();     // 向 [this+16] 写入 4 字节

对应 libQt6Gui 中的一条指令：

    str  w19, [x20, #16]        // 编码 B9 00 12 93（小端 93 12 00 b9）

对按 **Qt 6.4** 头文件编译的插件而言，偏移 16 是**插件自己的第一个成员**：

    class QFcitxPlatformInputContext : public QPlatformInputContext {
        ...
        FcitxWatcher *m_watcher;      // ← 偏移 16
    };

于是 Qt 在首次事件循环迭代时把 `m_watcher` 的低 32 位清零。国产系统上 python3 通常
**不是 PIE**，堆位于低位地址（实测 `0x3758xxxx`），低位清零后指针**恰好变成 NULL**：

* 运行期：`createICData()` 把空的 `m_watcher` 传给 `FcitxInputContextProxy`，
  出现 `QObject::connect: Cannot connect (nullptr)::availabilityChanged(bool) ...`，
  输入上下文建不起来 → **中文输入不可用**；
* 退出期：`~QFcitxPlatformInputContext()` 调用 `m_watcher->unwatch()` → 空指针解引用
  → `fault_addr=0x63`（即 `ldrb w0,[x0,#99]` 在 x0=NULL 时）→ 进程崩溃（SIGSEGV/SIGABRT）。

## 为什么可以安全地把该指令改成 NOP

* `m_inputDirection` 在 xcb 平台下**没有任何读取者**：`emitInputDirectionChanged()`
  只被 `libQt6WaylandClient.so.6` 引用（已在包内全库扫描确认），而 Qt 6.7 的
  `QPlatformInputContext::inputDirection()` 实现是 `QLocale::system().textDirection()`，
  并不读取该成员。
* 因此跳过这次赋值，对 xcb 下的输入法行为**没有任何可观察影响**，只是不再由 Qt 侧
  去覆盖（对 6.7 插件而言是它自己的成员，本来就无人读取）。

## 定位方式（不依赖硬编码偏移）

1. 从 `libQt6Gui.so.6` 的动态符号表取 `_ZNK21QPlatformInputContext14inputDirectionEv`
   与 `_ZN21QPlatformInputContextC1Ev` 两个符号地址，得到一个区间；
2. 在该区间内按字节搜索 `str w19,[x20,#16]` 的机器码 `93 12 00 b9`；
3. 要求**恰好命中一次**，否则拒绝修改；
4. 改写为 `nop`（`1f 20 03 d5`），并回读校验。

用法：
    python3 tools/patch_qt6_inputcontext.py <libQt6Gui.so.6> [输出文件]
    python3 tools/patch_qt6_inputcontext.py --check <libQt6Gui.so.6>
"""

import os
import re
import struct
import subprocess
import sys

SYM_BEFORE = "_ZNK21QPlatformInputContext14inputDirectionEv"
SYM_AFTER = "_ZN21QPlatformInputContextC1Ev"

# str w19, [x20, #16]  /  nop
INSN_STR = bytes.fromhex("931200b9")
INSN_NOP = bytes.fromhex("1f2003d5")


def _dynsym_addr(lib, name):
    """从动态符号表读取符号值（地址）。"""
    out = subprocess.run(["readelf", "--dyn-syms", "--wide", lib],
                         capture_output=True, text=True).stdout
    for line in out.splitlines():
        m = re.match(r"\s*\d+:\s+([0-9a-f]+)\s+\d+\s+\w+\s+\w+\s+\w+\s+\w+\s+(\S+)", line)
        if m and m.group(2).split("@")[0] == name:
            return int(m.group(1), 16)
    return None


def _vaddr_to_offset(lib, vaddr):
    """用程序头把虚拟地址换算成文件偏移。"""
    out = subprocess.run(["readelf", "-lW", lib], capture_output=True, text=True).stdout
    for line in out.splitlines():
        m = re.match(r"\s*LOAD\s+(0x[0-9a-f]+)\s+(0x[0-9a-f]+)\s+(0x[0-9a-f]+)\s+"
                     r"(0x[0-9a-f]+)\s+(0x[0-9a-f]+)", line)
        if not m:
            continue
        offset, vaddr_, _paddr, filesz, _memsz = (int(x, 16) for x in m.groups())
        if vaddr_ <= vaddr < vaddr_ + filesz:
            return offset + (vaddr - vaddr_)
    raise ValueError("无法把地址 0x%x 换算为文件偏移" % vaddr)


def _find_target(lib):
    """返回 (文件偏移, 已打补丁?) —— 定位那条 str w19,[x20,#16]。"""
    a = _dynsym_addr(lib, SYM_BEFORE)
    b = _dynsym_addr(lib, SYM_AFTER)
    if a is None or b is None:
        raise ValueError("缺少定位所需的符号（%s / %s）" % (SYM_BEFORE, SYM_AFTER))
    if b <= a:
        raise ValueError("符号顺序异常：inputDirection=0x%x, ctor=0x%x" % (a, b))
    off_a = _vaddr_to_offset(lib, a)
    off_b = _vaddr_to_offset(lib, b)
    with open(lib, "rb") as f:
        data = f.read()
    seg = data[off_a:off_b]
    hits = [i for i in range(len(seg) - 3) if seg[i:i + 4] == INSN_STR]
    if len(hits) != 1:
        raise ValueError("区间内 str w19,[x20,#16] 命中 %d 次（应为 1）" % len(hits))
    return off_a + hits[0], False


def check(lib):
    """判断是否已打补丁。返回 (状态字符串, 是否已修补)。"""
    a = _dynsym_addr(lib, SYM_BEFORE)
    b = _dynsym_addr(lib, SYM_AFTER)
    if a is None or b is None:
        return ("无法定位符号（可能不是 Qt6 的 libQt6Gui）", False)
    off_a, off_b = _vaddr_to_offset(lib, a), _vaddr_to_offset(lib, b)
    with open(lib, "rb") as f:
        data = f.read()
    seg = data[off_a:off_b]
    n_str = sum(1 for i in range(len(seg) - 3) if seg[i:i + 4] == INSN_STR)
    n_nop = sum(1 for i in range(len(seg) - 3) if seg[i:i + 4] == INSN_NOP)
    if n_str == 1:
        return ("未打补丁（仍存在 m_inputDirection 写入）", False)
    if n_str == 0 and n_nop >= 1:
        return ("已打补丁（写入指令已 NOP）", True)
    return ("状态不明：str=%d, nop=%d" % (n_str, n_nop), False)


def patch(lib, dst=None):
    off, _ = _find_target(lib)
    with open(lib, "rb") as f:
        data = bytearray(f.read())
    before = bytes(data[off:off + 4])
    data[off:off + 4] = INSN_NOP
    if bytes(data[off:off + 4]) != INSN_NOP:
        raise RuntimeError("写入校验失败")
    out = dst or lib
    with open(out, "wb") as f:
        f.write(data)
    return {"offset": off, "before": before.hex(), "after": INSN_NOP.hex(), "out": out}


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 2
    if args[0] == "--check":
        if len(args) < 2:
            print("用法：--check <libQt6Gui.so.6>")
            return 2
        state, done = check(args[1])
        print("状态:", state)
        print("RESULT:", "已修补" if done else "未修补")
        return 0 if done else 1
    src = args[0]
    dst = args[1] if len(args) > 1 else None
    info = patch(src, dst)
    print("已处理:", info["out"])
    print("  偏移 0x%x：%s → %s（str → nop）" % (info["offset"], info["before"], info["after"]))
    print("RESULT: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
