#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""修复"由缺陷版本产出的 .docx"：把误落在错误命名空间的关系条目搬回 OPC 包命名空间。

## 背景

旧版 `app/core/assembler.py` 的 `_patch_package()` 在补写图片/OLE 关系时误用了
`officeDocument/2006/relationships`（那是 `r:id` 属性与关系 `Type` 取值的前缀命名空间），
而 `word/_rels/document.xml.rels` 部件本身属于 OPC 的 `package/2006/relationships`。
产物里因此出现：

    <ns0:Relationship xmlns:ns0="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                      Id="rId1001" Type=".../image" Target="media/asm_1001.wmf"/>

命名空间不匹配时，OPC 解析器会把这些条目当作**外来元素忽略**：
`document.xml` 里的 `r:id` 于是悬空 —— 严格的解析器（Microsoft Word、LibreOffice）
会丢失图片/嵌入对象，甚至判定文件损坏；WPS 较宽容，可能看不出问题。
（症状也可能是"打不开 / source file could not be loaded"。）

本工具把这类条目原地改到正确命名空间，其余部件**逐字节保持不动**。

## 用法

    # 逐个文件（默认输出 <原名>_修复.docx，不改动原文件）
    python tools/fix_docx_rels_namespace.py 测试2/汇编结果（测试2）.docx

    # 整个目录（递归，跳过已修复的）
    python tools/fix_docx_rels_namespace.py 测试2 测试3

    # 直接覆盖原文件（会先备份为 <原名>.bak.docx）
    python tools/fix_docx_rels_namespace.py --in-place 某文件.docx

注意：本工具只修"关系命名空间"这一类缺陷。若文档还有别的问题（缺部件、
Content_Types 不全等），请用修复后的程序重新汇编一遍更稳妥。
"""
import argparse
import os
import shutil
import sys
import zipfile

from lxml import etree

PKG_R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NSDECL = 'xmlns="%s"' % PKG_R_NS


def _repair_rels_bytes(data):
    """就地修复一个 .rels 部件。返回 (新字节, 修复条目数)；无需修复时返回 (None, 0)。"""
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return None, 0
    if etree.QName(root).localname != "Relationships":
        return None, 0
    bad = [el for el in root
           if etree.QName(el).localname == "Relationship"
           and etree.QName(el).namespace != PKG_R_NS]
    if not bad:
        return None, 0
    for el in bad:
        # 直接改 tag：保留位置与全部属性，命名空间归入根元素的默认命名空间
        el.tag = "{%s}Relationship" % PKG_R_NS
    out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return out, len(bad)


def repair_file(path, out_path=None, in_place=False):
    """修复单个 docx。返回 (输出路径, 修复条目数)。修复条目数为 0 表示文件本就正常。"""
    with zipfile.ZipFile(path, "r") as z:
        infos = z.infolist()
        parts = {i.filename: z.read(i.filename) for i in infos}

    total = 0
    for name in list(parts):
        if not name.endswith(".rels"):
            continue
        new, n = _repair_rels_bytes(parts[name])
        if n:
            parts[name] = new
            total += n
            print("    %s：修复 %d 条关系" % (name, n))

    if not total:
        return path, 0

    if in_place:
        bak = path + ".bak.docx"
        shutil.copy2(path, bak)
        target = path
        print("    原文件已备份 -> %s" % os.path.basename(bak))
    else:
        target = out_path or (os.path.splitext(path)[0] + "_修复.docx")

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i.filename, parts[i.filename])
    return target, total


def verify(path):
    """轻量校验：关系条目命名空间是否全部正确、document.xml 引用是否都有解。"""
    import re
    import posixpath
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        raw = z.read("word/_rels/document.xml.rels").decode("utf-8")
        doc = z.read("word/document.xml").decode("utf-8", "replace")
    root = etree.fromstring(raw.encode("utf-8"))
    bad = [etree.QName(el).namespace for el in root
           if etree.QName(el).namespace != PKG_R_NS]
    valid = {el.get("Id") for el in root
             if etree.QName(el).namespace == PKG_R_NS}
    refs = set(re.findall(r'r:(?:id|embed|link)="([^"]+)"', doc))
    dangling = sorted(refs - valid)
    missing = []
    for el in root:
        if etree.QName(el).namespace != PKG_R_NS:
            continue
        tgt = (el.get("Target") or "").replace("\\", "/")
        if not tgt or tgt.startswith(("http:", "https:", "file:", "#")):
            continue
        part = _resolve_part("word/_rels/document.xml.rels", tgt)
        if part not in names:
            missing.append(part)
    return len(bad), dangling, missing


def _resolve_part(rels_name, target):
    """把关系 Target 解析为包内部件名。

    parts 目录 = .rels 文件所在目录的上一级（word/_rels/document.xml.rels -> word/），
    Target 相对它解析；`../`、`/` 开头都要按 posix 规则归一化。
    注意：不能简单地把 `..` 段丢弃——`../customXml/item1.xml` 指向 `customXml/item1.xml`，
    而 `word/customXml/item1.xml` 是另一个（不存在的）位置。
    """
    import posixpath
    base = posixpath.dirname(posixpath.dirname(rels_name))  # word/
    t = target.replace("\\", "/")
    if t.startswith("/"):
        return posixpath.normpath(t.lstrip("/"))
    return posixpath.normpath(posixpath.join(base, t))


def _iter_docx(paths):
    for p in paths:
        if os.path.isdir(p):
            for dirpath, _dirnames, filenames in os.walk(p):
                for fn in sorted(filenames):
                    if fn.lower().endswith(".docx") and not fn.startswith("~$"):
                        yield os.path.join(dirpath, fn)
        else:
            yield p


def main():
    ap = argparse.ArgumentParser(description="修复 docx 中误落命名空间的关系条目")
    ap.add_argument("paths", nargs="+", help="待修复的 .docx 文件或目录")
    ap.add_argument("--in-place", action="store_true",
                    help="直接覆盖原文件（会先备份为 <原名>.bak.docx）")
    ap.add_argument("--out-suffix", default="_修复",
                    help="非 in-place 时的输出后缀，默认 '_修复'")
    opts = ap.parse_args()

    files = [f for f in _iter_docx(opts.paths) if os.path.isfile(f)]
    if not files:
        print("未找到 .docx 文件")
        return 1

    n_fix = 0
    n_ok = 0
    n_bad = 0
    for f in files:
        print("· %s" % os.path.relpath(f))
        out = None if opts.in_place else (
            os.path.splitext(f)[0] + opts.out_suffix + ".docx")
        if out and os.path.exists(out):
            print("    已存在同名输出，跳过（先删除或改用 --in-place）")
            continue
        try:
            target, n = repair_file(f, out, opts.in_place)
        except zipfile.BadZipFile:
            print("    [跳过] 不是有效的 docx/zip")
            n_bad += 1
            continue
        if not n:
            print("    无需修复（关系命名空间正常）")
            n_ok += 1
            continue
        bad, dangling, missing = verify(target)
        ok = (bad == 0 and not dangling and not missing)
        print("    输出 -> %s" % os.path.relpath(target))
        print("    复核：错误命名空间 %d；悬空引用 %s；缺失部件 %s  =>  %s"
              % (bad, dangling or "无", missing or "无", "通过 ✓" if ok else "仍异常 ✗"))
        n_fix += 1
        if not ok:
            n_bad += 1

    print("\n汇总：修复 %d 个，本就正常 %d 个，异常 %d 个" % (n_fix, n_ok, n_bad))
    return 0


if __name__ == "__main__":
    sys.exit(main())
