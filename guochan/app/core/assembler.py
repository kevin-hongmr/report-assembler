# -*- coding: utf-8 -*-
"""
assembler.py —— 汇编与格式引擎（核心）
把若干已转为 .docx 的源文档按《汇编格式模板》规范汇编成一份 docx。
严格应用用户 11 条格式要求：
  1) 支持 doc/docx/wps（wps/doc 由 converter 预转为 docx，本模块只吃 docx）
  2) 封面 + 目录 + 正文
  3) 每份文档标题：居中、方正小标宋简体、二号(22pt)
  4) 正文：首行缩进2字符、三号(16pt)、中文仿宋_GB2312、西文宋体
  5) 一级标题：黑体、三号、首行缩进2字符
  6) 二级标题：楷体_GB2312、三号、首行缩进2字符
  7) 整份行距固定 28.5 磅
  8) 每份文档单位写在标题下：楷体_GB2312、三号、居中
  9) 删除原文档落款（单位+日期）
  10) 页脚页码 —N—，双页翻面（奇偶不同页脚），正文首页为第1页
  11) A4，左右2.7cm、上下3cm
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
import os
import copy
import zipfile
import hashlib
from lxml import etree

# 关系（relationships）命名空间，用于图片/编号等外部引用
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# .rels 根元素 <Relationships> 及其子元素 <Relationship> 所用的命名空间。
# 注意：这与 R_NS 不同——R_NS 是 r:id 属性与关系 Type 取值的前缀命名空间，
# 而 .rels 部件本身属于 OPC 的 package 命名空间。新建 <Relationship> 时若误用
# R_NS，元素会落到错误命名空间，OPC 解析器（LibreOffice/Word）会将其视为外来
# 元素忽略，导致 document.xml 的 r:id 悬空、文件被判定为无法加载（预览失败）。
PKG_R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"

# ---------- 字体常量 ----------
FONT_TITLE   = "方正小标宋简体"   # 封面大标题 / 每份材料标题 / 目录标题 / 封面日期
FONT_UNIT    = "楷体_GB2312"      # 单位行 / 二级标题
FONT_H1      = "黑体"             # 一级标题
FONT_H2      = "楷体_GB2312"      # 二级标题
FONT_BODY_EA = "仿宋_GB2312"      # 正文中文
FONT_BODY_ASCII = "宋体"          # 正文西文（用户规格第4条）
FONT_TOC     = "仿宋_GB2312"      # 目录条目

# ---------- 字号（磅）----------
SIZE_BODY    = 16   # 三号
SIZE_TITLE   = 22   # 二号
SIZE_TOC     = 16   # 三号
SIZE_DATE    = 16   # 三号

LINE_PT = 28.5      # 固定行距（磅）

TWIPS_PER_CM = 567
MARGIN_LR_CM = 2.7
MARGIN_TB_CM = 3.0

# 落款识别
_DATE_RE = re.compile(r"(20\d{2}年\d{1,2}月\d{1,2}日|20\d{2}[-/]\d{1,2}[-/]\d{1,2}|"
                      r"\d{4}年\d{1,2}月|\d{1,2}年\d{1,2}月\d{1,2}日|"
                      r"二[〇零一二三四五六七八九十]+年[一二三四五六七八九十]+月[一二三四五六七八九十]+日)")
_UNIT_HINT = ("局", "委", "政府", "办公室", "公司", "集团", "院", "学校", "中心",
              "部", "处", "所", "会", "队", "站", "厅", "署", "区委", "区委")


# ============================================================
# 底层样式辅助
# ============================================================
_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
              "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
              "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
              "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
              "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
              "eastAsianLayout", "specVanish", "oMath"]


def _insert_rpr_child(rPr, el):
    """按 OOXML CT_RPr 规范顺序插入子元素。

    w:spacing（字符间距）必须排在 w:color 之后、w:sz 之前；w:rFonts 也须靠前。
    顺序违规会让 Word 判定文档损坏并拒绝打开。
    """
    tag = etree.QName(el).localname
    idx = _RPR_ORDER.index(tag) if tag in _RPR_ORDER else len(_RPR_ORDER)
    for child in rPr:
        ct = etree.QName(child).localname
        if ct in _RPR_ORDER and _RPR_ORDER.index(ct) > idx:
            child.addprevious(el)
            return
    rPr.append(el)


def _set_run_fonts(run, eastasia=None, ascii_font=None, hansi=None, size=None, bold=None):
    """设置 run 的字体（含中文 eastAsia、西文 ascii/hAnsi、复杂文种 cs）。

    兼容性要点：OOXML 的 run 属性存在『主属性 / 复杂文种(Cs)』成对的情况，
    源文档往往两者都写了、且可能不一致。若只改主属性、残留 Cs 属性，WPS 渲染中文时
    会按 Cs 属性显示（如 szCs=44 → 整段二号）。因此凡是主动设置的属性，其 Cs 变体一并同步：
      - 字体 rFonts 的 @cs（复杂文种字体）＝ 西文字体 ascii
      - 字号 w:szCs ＝ w:sz
      - 加粗 w:bCs ＝ w:b
    """
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        _insert_rpr_child(rPr, rFonts)
    if eastasia:
        rFonts.set(qn("w:eastAsia"), eastasia)
    if ascii_font:
        rFonts.set(qn("w:ascii"), ascii_font)
        # 复杂文种字体同步为中文字体（源文档 @cs 残留如 黑体/方正小标宋 会被清掉；
        # 实测 WPS 中文用 eastAsia、但 size 又按 szCs，为稳妥 cs 跟随中文字体）
        rFonts.set(qn("w:cs"), eastasia or ascii_font)
    if hansi:
        rFonts.set(qn("w:hAnsi"), hansi)
    if size is not None:
        run.font.size = Pt(size)
        # 关键：同时写 w:szCs。源文档正文常带 szCs=44（二号），python-docx 只写 w:sz
        # 会残留旧 szCs → WPS 渲染中文时按 szCs 显示，出现"整段变二号"的 bug。
        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            _insert_rpr_child(rPr, szCs)
        szCs.set(qn("w:val"), str(int(Pt(size).pt * 2)))
    # 加粗：强制时设 b 并同步 bCs；保留原文时（bold=None，正文）也把 bCs 对齐到当前 b，
    # 避免源文档 b/bCs 不一致（源常有 b=加粗、bCs=0）导致 WPS 对复杂文种渲染不一致。
    if bold is None:
        b = rPr.find(qn("w:b"))
        if b is not None:
            bold = (b.get(qn("w:val")) not in ("0", "false"))
    if bold is not None:
        run.font.bold = bold
        bCs = rPr.find(qn("w:bCs"))
        if bCs is None:
            bCs = OxmlElement("w:bCs")
            _insert_rpr_child(rPr, bCs)
        bCs.set(qn("w:val"), "1" if bold else "0")
    # 字符间距：标准（0）。源文档的 run 可能带非标准字距，需显式归零
    sp = rPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        _insert_rpr_child(rPr, sp)
    sp.set(qn("w:val"), "0")


# 材料标题字符样式 id：标题字体/字号/加粗走样式而非直接格式。
# 关键：WPS『更新目录域』会把标题 run 上的**直接**字体名复制进目录条目
# （实测条目 rPr 出现 rFonts=方正小标宋简体，导致目录内容不是正文字体），
# 但**样式级**格式不会传播——故标题用字符样式，目录条目即保持 toc 样式（仿宋三号）。
ASM_TITLE_STYLE = "AsmTitle"


def _add_title_char_style(document):
    """定义材料标题字符样式（方正小标宋简体、二号、**不加粗**、字符间距标准）。"""
    styles_el = document.styles.element
    for st in styles_el.findall(qn("w:style")):
        if st.get(qn("w:styleId")) == ASM_TITLE_STYLE:
            styles_el.remove(st)
    st = OxmlElement("w:style")
    st.set(qn("w:type"), "character")
    st.set(qn("w:styleId"), ASM_TITLE_STYLE)
    nm = OxmlElement("w:name"); nm.set(qn("w:val"), ASM_TITLE_STYLE); st.append(nm)
    rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT_TITLE)
    rPr.append(rf)
    # 规范：材料标题不加粗——显式置 0（防止基础样式加粗生效）
    b = OxmlElement("w:b"); b.set(qn("w:val"), "0"); rPr.append(b)
    bcs = OxmlElement("w:bCs"); bcs.set(qn("w:val"), "0"); rPr.append(bcs)
    sp = OxmlElement("w:spacing"); sp.set(qn("w:val"), "0"); rPr.append(sp)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(SIZE_TITLE * 2)); rPr.append(sz)
    szcs = OxmlElement("w:szCs"); szcs.set(qn("w:val"), str(SIZE_TITLE * 2)); rPr.append(szcs)
    st.append(rPr)
    styles_el.append(st)


def _apply_title_style_run(run):
    """给材料标题 run 挂字符样式（不写直接 rPr 字体，避免更新目录域时字体被复制进条目）。"""
    rPr = run._element.get_or_add_rPr()
    rs = rPr.find(qn("w:rStyle"))
    if rs is None:
        rs = OxmlElement("w:rStyle")
        rPr.insert(0, rs)
    rs.set(qn("w:val"), ASM_TITLE_STYLE)


def _set_first_line_indent(p, chars=2, size_pt=None):
    """首行缩进 N 个字符。**两个属性必须同时写**，缺一不可。

    * ``w:firstLineChars = chars*100`` —— 东亚版式「按字符缩进」，WPS / Word 认，
      随字号自动换算，是成品的权威值；
    * ``w:firstLine = chars × 字号(pt) × 20`` —— 等值 twips。
      **LibreOffice 完全忽略 firstLineChars**（实测 LO 6.4：只写 Chars 时首行缩进
      为 0；只写 firstLine 时得到 32 磅 = 2 个 16 磅字符）。预览/PDF 正是走
      LibreOffice，只写 Chars 就会出现「成品正常、预览里首行不缩进」。
      两者在同一段落上数值等价，谁被采纳结果都一样，故并存安全。

    还必须删除 ``w:hanging`` / ``w:hangingChars``（不是置 0，是删掉）：
      * ``w:hangingChars`` 哪怕写 0，WPS 也会忽略 firstLineChars；
      * ``w:hanging`` 哪怕写 0，**LibreOffice 也会忽略 firstLine**（两者映射到
        同一个内部属性，hanging 优先）。实测：firstLine=640 + hanging=0 → 不缩进。
    悬挂缩进改由「不写 hanging + 显式写 firstLine」抵消样式/编号里的悬挂值。
    """
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        _insert_ppr_child(pPr, ind)
    for a in ("w:firstLine", "w:hanging", "w:hangingChars"):
        if ind.get(qn(a)) is not None:
            del ind.attrib[qn(a)]
    try:
        chars = int(chars or 0)
    except (TypeError, ValueError):
        chars = 0
    try:
        size = float(size_pt) if size_pt else float(SIZE_BODY)
    except (TypeError, ValueError):
        size = float(SIZE_BODY)
    if size <= 0:
        size = float(SIZE_BODY)
    ind.set(qn("w:firstLineChars"), str(chars * 100))
    ind.set(qn("w:firstLine"), str(int(round(chars * size * 20))))


def _set_exact_line(p, pt=LINE_PT):
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(pt)


def _set_line(p, mode="exact", value=LINE_PT):
    """设置段落行距：mode='exact' 为固定值（磅）；mode='multiple' 为行距倍数。

    OOXML 映射：固定值 -> w:line=<磅*20> lineRule=exact；
    行距倍数 -> w:line=<倍数*240> lineRule=auto（1 倍=240）。
    """
    if mode == "multiple":
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p.paragraph_format.line_spacing = float(value)
    else:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(float(value))


def _add_para(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT,
              eastasia=None, ascii_font=None, hansi=None,
              size=None, bold=None, indent_chars=None, line_pt=LINE_PT):
    """通用：添加一个按规格排版的段落。"""
    p = doc.add_paragraph()
    p.alignment = align
    if indent_chars is not None:
        _set_first_line_indent(p, indent_chars, size)
    _set_exact_line(p, line_pt)
    if text:
        run = p.add_run(text)
        _set_run_fonts(run, eastasia=eastasia, ascii_font=ascii_font,
                       hansi=hansi or ascii_font, size=size, bold=bold)
    # 缩进显式置 0（首行按入参，左右/悬挂恒 0），避免继承样式带来的缩进
    _set_first_line_indent(p, indent_chars if indent_chars is not None else 0, size)
    _zero_horizontal_indents(p)
    return p


def _level_fonts(level, sizes=None):
    """返回某层级的中文字体/西文字体/字号/默认加粗。

    sizes: {"body": 磅, "h1": 磅, "h2": 磅}（可选，用户界面可调字号；
    缺省用默认三号）。标题/单位/封面等层级不接受 sizes 调整。
    默认加粗含义：非 None 时强制该层级所有 run 的加粗（忽略原文加粗）；
    为 None 时（正文）保留原文 run 级加粗。
    """
    sizes = sizes or {}
    sz_body = float(sizes.get("body", SIZE_BODY))
    sz_h1 = float(sizes.get("h1", SIZE_BODY))
    sz_h2 = float(sizes.get("h2", SIZE_BODY))
    if level == "title":
        return FONT_TITLE, FONT_TITLE, SIZE_TITLE, False   # 材料标题：不加重粗
    if level == "unit":
        return FONT_UNIT, FONT_UNIT, SIZE_BODY, False
    if level == "h1":
        return FONT_H1, FONT_H1, sz_h1, False              # 一级标题：黑体、不加粗
    if level == "h2":
        return FONT_H2, FONT_H2, sz_h2, True               # 二级标题：楷体_GB2312、加粗
    # body
    return FONT_BODY_EA, FONT_BODY_ASCII, sz_body, None    # 正文：仿宋/宋体，保留原文加粗


def _add_para_runs(doc, runs, level, align=WD_ALIGN_PARAGRAPH.LEFT,
                   indent_chars=None, line_pt=LINE_PT, line_mode="exact",
                   sizes=None, space_before=None, space_after=None):
    """按层级套用字体，但保留各 run 的原始加粗；可选段前/段后间距。

    line_mode='exact' 时 line_pt 为固定磅值；'multiple' 时 line_pt 为行距倍数。
    sizes 为可调字号（{body,h1,h2: 磅}），仅 body/h1/h2 层级生效。
    """
    ea, ascii_font, size, bold_default = _level_fonts(level, sizes)
    p = doc.add_paragraph()
    p.alignment = align
    if indent_chars is not None:
        _set_first_line_indent(p, indent_chars, size)
    _set_line(p, line_mode, line_pt)
    if space_before is not None:
        p.paragraph_format.space_before = space_before
    if space_after is not None:
        p.paragraph_format.space_after = space_after
    for (txt, bold) in runs:
        if not txt:
            continue
        run = p.add_run(txt)
        if level == "title":
            # 材料标题：字体/字号/加粗走字符样式（见 ASM_TITLE_STYLE 注释——
            # 直接格式会被 WPS 更新目录域时复制进目录条目）
            _apply_title_style_run(run)
            continue
        b = bold if bold_default is None else bold_default
        _set_run_fonts(run, eastasia=ea, ascii_font=ascii_font,
                       hansi=ascii_font, size=size, bold=b)
    # 缩进显式置 0（首行按入参，左右/悬挂恒 0），避免继承样式带来的缩进
    _set_first_line_indent(p, indent_chars if indent_chars is not None else 0, size)
    _zero_horizontal_indents(p)
    return p


def _split_runs_at_period(runs):
    """把一个段落的 runs 在第一个全角句号『。』处分割为 (标题部分, 正文部分)。

    标题部分包含句号本身；若段落中没有句号，则整体作为标题部分返回。
    runs: list of (text, bold)
    """
    full = "".join(t for t, _ in runs)
    cut = full.find("。")
    if cut == -1:
        return list(runs), []
    cut_end = cut + 1  # 句号归标题部分
    heading, body = [], []
    pos = 0
    for (t, b) in runs:
        if not t:
            continue
        if pos >= cut_end:
            body.append((t, b))
        elif pos + len(t) <= cut_end:
            heading.append((t, b))
        else:
            s = cut_end - pos
            heading.append((t[:s], b))
            body.append((t[s:], b))
        pos += len(t)
    return heading, body


def _add_heading_para(doc, runs, level, indent_chars=2):
    """添加一级/二级标题段落，统一段前段后间距为 0。

    - 一级标题：整体黑体、不加粗。
    - 二级标题：从『（一）』等标识到该句第一个句号（含句号）使用楷体_GB2312 加粗；
      句号之后的文字视为正文，使用 仿宋_GB2312（中文）/ 宋体（西文），并保留原始加粗。
    """
    ea, ascii_font, size, bold_default = _level_fonts(level)
    if level == "h2":
        heading_runs, body_runs = _split_runs_at_period(runs)
    else:
        heading_runs, body_runs = runs, []
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if indent_chars:
        _set_first_line_indent(p, indent_chars, size)
    _set_exact_line(p, LINE_PT)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    # 标题部分
    for (txt, b) in heading_runs:
        if not txt:
            continue
        r = p.add_run(txt)
        _set_run_fonts(r, eastasia=ea, ascii_font=ascii_font, hansi=ascii_font,
                       size=size, bold=bold_default)
    # 二级标题句号后的正文部分
    if body_runs:
        ea_b, ascii_b, size_b, _ = _level_fonts("body")
        for (txt, b) in body_runs:
            if not txt:
                continue
            r = p.add_run(txt)
            _set_run_fonts(r, eastasia=ea_b, ascii_font=ascii_b, hansi=ascii_b,
                           size=size_b, bold=b)
    return p


def _add_page_break(doc):
    """插入一个分页符段落。"""
    p = doc.add_paragraph()
    r = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r._r.append(br)


def _add_page_field(paragraph):
    """在段落中插入 — PAGE — （页码域）。"""
    paragraph.add_run("—")
    r = paragraph.add_run()
    fldBegin = OxmlElement("w:fldChar"); fldBegin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " PAGE "
    fldEnd = OxmlElement("w:fldChar"); fldEnd.set(qn("w:fldCharType"), "end")
    r._r.append(fldBegin); r._r.append(instr); r._r.append(fldEnd)
    paragraph.add_run("—")


def _set_outline_level(p, level):
    """给段落设置大纲级别（用于自动目录域捕获；level=1 -> val 0）。"""
    pPr = p._p.get_or_add_pPr()
    ol = pPr.find(qn("w:outlineLvl"))
    if ol is None:
        ol = OxmlElement("w:outlineLvl")
        _insert_ppr_child(pPr, ol)
    ol.set(qn("w:val"), str(level - 1))


def _strip_outline_level(p):
    """把段落标记为正文（大纲级别 9），避免其混入目录。

    目录域按大纲级别抓取条目，只有材料标题应保留级别。克隆段落若只删除
    直接 w:outlineLvl，仍会继承其 pStyle（如『标题 1』样式自带 outlineLvl=0）
    的大纲级别，更新目录域后照样被收进目录（实测：测试3 固投汇报的一级标题
    套用了 heading 1 样式）。因此这里显式置 outlineLvl=9（正文），覆盖样式
    继承，确保只有材料标题进入目录。
    """
    pPr = p._p.get_or_add_pPr()
    ol = pPr.find(qn("w:outlineLvl"))
    if ol is None:
        ol = OxmlElement("w:outlineLvl")
        _insert_ppr_child(pPr, ol)
    ol.set(qn("w:val"), "9")


_SALUTE_KEYWORDS = ("尊敬的", "各位领导", "各位", "同志们", "同志", "在座",
                    "先生", "女士", "领导：", "参会")
def _is_salutation(text):
    """判断是否为抬头/称呼行（如「尊敬的陈区长、各位领导：」），此类行不缩进。"""
    t = (text or "").strip()
    if not t:
        return False
    return any(k in t for k in _SALUTE_KEYWORDS)


def _add_toc_field(document):
    """插入 WPS/Word 自带「引用→目录」域（自动生成页码）。打开时更新域即可。"""
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_exact_line(p, LINE_PT)
    r = p.add_run("目  录")
    _set_run_fonts(r, eastasia=FONT_TITLE, ascii_font=FONT_TITLE, hansi=FONT_TITLE,
                   size=SIZE_TITLE, bold=True)
    # TOC 域：仅按大纲级别 1 捕获各材料标题（\o "1-1"）。
    # 旧版 \o "1-3" 会把一级/二级标题（程序设过级别）及源文档自带级别的正文
    # 一并收进目录；现只给材料标题保留大纲级别，其余段落全部剥离。
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_exact_line(p, LINE_PT)

    def _fld(typ):
        run = p.add_run()
        fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), typ)
        run._r.append(fc)
        return run
    _fld("begin")
    r = p.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-1" \\h \\u '
    r._r.append(instr)
    _fld("separate")
    r = p.add_run()
    r.text = "（在 WPS/Word 中打开后，右键『更新域』生成目录页码）"
    _fld("end")


def _add_toc_styles(document, toc_size=None, line_mode="exact", line_value=LINE_PT):
    """定义目录条目样式（toc 1 ~ toc 3），使『更新域』生成的目录内容与正文格式一致：
    中文仿宋_GB2312 / 西文宋体、三号(16pt)、**无首行缩进**、行距固定 28.5 磅、
    段前段后 0、两端对齐、字符间距标准。WPS/Word 更新目录域时按 toc N 样式生成条目。

    toc_size（磅）/ line_mode（exact|multiple）/ line_value（磅值或倍数）为界面可调项，
    仅作用于目录条目；「目 录」标题与材料标题不受影响。
    """
    toc_size = float(toc_size if toc_size is not None else SIZE_TOC)
    if line_mode == "multiple":
        line_val = str(int(float(line_value) * 240))
        line_rule = "auto"
    else:
        line_val = str(int(float(line_value) * 20))
        line_rule = "exact"
    sz_val = str(int(toc_size * 2))
    styles_el = document.styles.element
    # 移除已有的 TOC1~3 定义（默认模板可能带潜在样式），再写入规范样式
    for st in styles_el.findall(qn("w:style")):
        if st.get(qn("w:styleId")) in ("TOC1", "TOC2", "TOC3"):
            styles_el.remove(st)
    for n in (1, 2, 3):
        st = OxmlElement("w:style")
        st.set(qn("w:type"), "paragraph")
        st.set(qn("w:styleId"), f"TOC{n}")
        nm = OxmlElement("w:name"); nm.set(qn("w:val"), f"toc {n}"); st.append(nm)
        based = OxmlElement("w:basedOn"); based.set(qn("w:val"), "Normal"); st.append(based)
        st.append(OxmlElement("w:unhideWhenUsed"))
        pPr = OxmlElement("w:pPr")
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:before"), "0"); sp.set(qn("w:after"), "0")
        sp.set(qn("w:line"), line_val); sp.set(qn("w:lineRule"), line_rule)
        _insert_ppr_child(pPr, sp)
        ind = OxmlElement("w:ind")
        # 目录条目不做首行缩进（用户要求）；左右/悬挂仍显式置 0 避免继承样式缩进
        for a in ("w:left", "w:leftChars", "w:right", "w:rightChars", "w:hanging"):
            ind.set(qn(a), "0")
        _insert_ppr_child(pPr, ind)
        jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "both")
        _insert_ppr_child(pPr, jc)
        st.append(pPr)
        rPr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        rf.set(qn("w:ascii"), "宋体"); rf.set(qn("w:hAnsi"), "宋体")
        rf.set(qn("w:eastAsia"), "仿宋_GB2312"); rf.set(qn("w:cs"), "宋体")
        rPr.append(rf)
        rsp = OxmlElement("w:spacing"); rsp.set(qn("w:val"), "0"); rPr.append(rsp)
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), sz_val); rPr.append(sz)
        szcs = OxmlElement("w:szCs"); szcs.set(qn("w:val"), sz_val); rPr.append(szcs)
        st.append(rPr)
        styles_el.append(st)


# ============================================================
# 源文档解析
# ============================================================
def _run_fonts_of(paragraph):
    """返回段落首个非空 run 的字体名（粗略用于标题识别）。"""
    for r in paragraph.runs:
        if r.text.strip():
            rPr = r._element.find(qn("w:rPr"))
            if rPr is not None:
                rf = rPr.find(qn("w:rFonts"))
                if rf is not None:
                    return (rf.get(qn("w:eastAsia")) or "", rf.get(qn("w:ascii")) or "")
    return ("", "")


def classify(text, ea_font=""):
    """把一段文字归类：h1 / h2 / body。文本模式优先（中文标题最可靠），字体兜底。"""
    t = text.strip()
    if not t:
        return "empty"
    # 文本模式优先（中文标题最可靠）
    # 一级标题：一、二、…（中文数字 + 顿号/句号/点）
    if re.match(r"^[一二三四五六七八九十百零]+[、．.]", t):
        return "h1"
    # 二级标题：（一）（二）…（中文数字加括号）
    if re.match(r"^（[一二三四五六七八九十百零]+）", t):
        return "h2"
    # 注意：阿拉伯数字编号（1. 2. (1) 等）不由 classify 判级（返回 body），
    #       块级分类里由 _H3_RE 识别为三级标题（保留原文格式），
    #       旧版曾把『1.加快完成…』误判为楷体二级标题
    # 字体兜底（仅当无文本模式命中时，作为较弱线索）
    if ea_font:
        if "黑体" in ea_font:
            return "h1"
        if "楷体" in ea_font:
            return "h2"
    return "body"


def extract_from_docx(path):
    """读取 docx，返回段落列表 [{text, level}]。level ∈ h1/h2/body/empty。"""
    doc = Document(path)
    out = []
    for p in doc.paragraphs:
        text = p.text
        ea, _ = _run_fonts_of(p)
        lvl = classify(text, ea)
        out.append({"text": text, "level": lvl})
    return out


# 标题结尾关键词（用于从文档开头识别材料标题）
_TITLE_END = ("汇报", "表态", "发言", "计划", "情况", "措施", "建议",
              "讲话", "总结", "报告", "交流", "介绍", "说明", "材料")
# 单位特征字（用于识别标题下的单位行）
_UNIT_HINT2 = ("局", "委", "政府", "委员会", "管委会", "管理", "公司", "集团",
              "院", "学校", "中心", "部", "处", "所", "会", "队", "站", "厅",
              "署", "区委", "办公室", "分局", "总队", "支队", "大队", "医院", "银行")


def _looks_date(t):
    """判断一段文字是否主要为日期（含 8 位数字、年月日、年月等）。"""
    if re.search(r"\d{8}", t):
        return True
    if re.search(r"20\d{2}年\d{1,2}月\d{1,2}日", t):
        return True
    if re.search(r"20\d{2}年\d{1,2}月", t):  # 也匹配「2026年8月*日」等日被遮蔽情形
        return True
    if re.search(r"\d{4}年\d{1,2}月\d{1,2}日", t):
        return True
    return False


def _looks_unit(t):
    """判断一段文字是否像单位名。"""
    return any(h in t for h in _UNIT_HINT2)


def _ea_of(p):
    """段落首个非空 run 的 eastAsia 字体。"""
    for r in p.runs:
        if r.text.strip():
            rPr = r._element.find(qn("w:rPr"))
            if rPr is not None:
                rf = rPr.find(qn("w:rFonts"))
                if rf is not None:
                    return rf.get(qn("w:eastAsia")) or ""
    return ""


def _has_numpr(p):
    pPr = p._p.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:numPr")) is not None


def _numid_of(p):
    pPr = p._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    np = pPr.find(qn("w:numPr"))
    if np is None:
        return None
    nid = np.find(qn("w:numId"))
    return nid.get(qn("w:val")) if nid is not None else None


# 三级编号模式：阿拉伯数字编号（1. / 1、/ （1）/ 1）/ ①-⑳ 等）开头的段落。
# 编号后紧跟数字的排除（如『1.5个百分点』『2024.』不算）。
_H3_RE = re.compile(
    r"^(?:\d{1,3}[．.、]|[（(]\d{1,3}[)）]|\d{1,3}[)）]|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])(?!\d)")
# 三级标题长度上限：仅限短标题行；更长的编号段落是正文列表项
# （如『1.大岭山产业园…：总投资…』，源里可能带异常字号，须按正文统一三号）
_H3_MAX_LEN = 30


def _classify_block(text, ea, has_numpr, numid):
    """块级分类（用于套字体），在 classify 文本/字体规则基础上补充：
    自动编号（numPr）但无手动『一、/（一）』前缀的短句（≤20字、不以。/数字/（结尾）
    视为一级标题，使其获得黑体并保留自动编号序号；
    阿拉伯数字编号（1. 1、（1） 1） ① 等）开头的短行（≤30字）视为三级标题——保留原文格式。
    """
    lvl = classify(text, ea)
    if lvl in ("h1", "h2"):
        return lvl
    if has_numpr:
        t = text.strip()
        if len(t) <= 20 and not t.endswith("。") and not re.match(r"^[0-9（(]", t):
            return "h1"
    t = text.strip()
    if len(t) <= _H3_MAX_LEN and _H3_RE.match(t):
        return "h3"
    return lvl  # body 或字体兜底


def _clean_signoff_blocks(blocks):
    """删除末尾落款：先去尾部空段，再删可能的日期段，最后删短单位段。"""
    while blocks and blocks[-1][0] == "p" and not blocks[-1][2].text.strip():
        blocks.pop()
    if not blocks:
        return blocks
    last = blocks[-1]
    if last[0] == "p" and _DATE_RE.search(last[2].text):
        blocks.pop()
        while blocks and blocks[-1][0] == "p" and not blocks[-1][2].text.strip():
            blocks.pop()
        if blocks and blocks[-1][0] == "p":
            lt = blocks[-1][2].text.strip()
            if len(lt) <= 20 and _looks_unit(lt):
                blocks.pop()
    while blocks and blocks[-1][0] == "p" and not blocks[-1][2].text.strip():
        blocks.pop()
    return blocks


def analyze(path):
    """从源 docx 解析出汇编所需的全部信息。

    返回：
      title / unit / date —— 从内容识别的标题、单位、日期
      dx    —— 源 docx 路径（克隆块、补媒体/编号时用）
      doc   —— 已打开的源 Document（保持引用，使块元素有效）
      blocks—— 正文有序块列表，每项：
               ("p",  <w:p>元素>, 段落对象) 或 ("tbl", <w:tbl>元素>, None)
               已排除 标题/单位/日期 段落，以及首尾空段；末尾落款已清理。
               表格原样保留（后续克隆时不重排格式）。
      paras —— 兼容字段（仅文本+级别），供旧逻辑/测试使用。

    标题识别关键修正：文档标题是『第一个非一级/二级标题模式的段落』，
    不再误把『一、…情况』这类一级标题当作文档标题；单位/日期仅限短行（≤25字），
    避免把含『区委』等字样的正文整段吞为『单位』。
    """
    doc = Document(path)
    body = doc.element.body
    paras = doc.paragraphs
    p_by_el = {}
    for p in paras:
        p_by_el[p._p] = p

    raw = []
    for p in paras:
        runs = [(r.text, bool(r.bold)) for r in p.runs]
        ea = _ea_of(p)
        raw.append({"p": p, "text": p.text, "align": p.alignment, "ea": ea, "runs": runs})

    # ---- 标题：第一个非 h1/h2 模式的非空段落 ----
    nonempty = [i for i, pr in enumerate(raw) if pr["text"].strip()]
    title = ""
    title_indices = []
    title_idx = -1
    for i in nonempty[:5]:
        if classify(raw[i]["text"], raw[i]["ea"]) not in ("h1", "h2"):
            title_idx = i
            title_indices = [i]
            title = raw[i]["text"].strip()
            break
    if title_idx < 0:  # 前 5 段全是标题模式：回退取首个非空段落
        title_idx = nonempty[0]
        title_indices = [title_idx]
        title = raw[title_idx]["text"].strip()

    # 合并后续居中短行（标题续行；遇带括号/单位特征的行停止）
    n = len(raw)
    j = title_idx + 1
    while j < n:
        t = raw[j]["text"].strip()
        if t == "":
            j += 1
            continue
        if (raw[j]["align"] == WD_ALIGN_PARAGRAPH.CENTER
                and len(t) <= 22 and not _looks_unit(t) and not t.startswith("（")):
            title_indices.append(j)
            title = title + "\n" + t
            j += 1
        else:
            break

    # ---- 标题行尾的括号单位：如『…税费收入分析  （税务局）』——
    # 剥离括号部分作为单位，标题净后由汇编规范格式（楷体居中单位行）输出。
    # 仅在括号内文字像单位名时剥离，避免误剥标题自身末尾的『（2025年上半年）』等注释。
    paren_unit = ""
    m = re.search(r"[（(]\s*([^（）()]+?)\s*[）)]\s*$", title)
    if m and _looks_unit(m.group(1)):
        paren_unit = m.group(1).strip()
        title = title[:m.start()].rstrip()

    # ---- 元信息（单位 / 日期）：仅限短行，避免吞掉正文 ----
    unit = ""
    date = ""
    meta_indices = []
    while j < n:
        t = raw[j]["text"].strip()
        if t == "":
            if meta_indices:
                break
            j += 1
            continue
        is_meta = ((raw[j]["align"] == WD_ALIGN_PARAGRAPH.CENTER
                    or _looks_unit(t) or _looks_date(t)) and len(t) <= 25)
        if not is_meta:
            break
        meta_indices.append(j)
        j += 1
    for m in meta_indices:
        m2 = raw[m]["text"].strip().strip("（）() ")
        if _looks_date(m2) and _looks_unit(m2):
            dm = re.search(r"(\d{8}|20\d{2}年\d{1,2}月\d{1,2}日|20\d{2}年\d{1,2}月|\d{4}年\d{1,2}月\d{1,2}日)", m2)
            if dm:
                date = dm.group(1)
                ut = m2.replace(dm.group(1), "").strip("（）() ")
                if _looks_unit(ut) and not unit:
                    unit = ut
            elif _looks_unit(m2) and not unit:
                unit = m2
        elif _looks_date(m2):
            date = m2
        elif _looks_unit(m2):
            if not unit:
                unit = m2

    # 若没有独立的单位行，但标题行尾剥出了括号单位，作为单位输出
    if not unit and paren_unit:
        unit = paren_unit

    # ---- 排除集合（标题/单位/日期段落对象）----
    excluded = set()
    for i in title_indices + meta_indices:
        if 0 <= i < n:
            excluded.add(raw[i]["p"])

    # ---- 正文 blocks（按 body 子元素顺序，含表格；排除标题/单位/日期与首尾空段）----
    blocks = []
    for child in body:
        if child.tag == qn("w:p"):
            p = p_by_el.get(child)
            if p is None:
                continue
            if p in excluded:
                continue
            blocks.append(("p", child, p))
        elif child.tag == qn("w:tbl"):
            blocks.append(("tbl", child, None))

    while blocks and blocks[0][0] == "p" and not blocks[0][2].text.strip():
        blocks.pop(0)
    while blocks and blocks[-1][0] == "p" and not blocks[-1][2].text.strip():
        blocks.pop()

    # ---- 清理末尾落款 ----
    blocks = _clean_signoff_blocks(blocks)

    # 兼容字段 paras（仅文本+级别）
    compat = []
    for kind, el, p in blocks:
        if kind == "p":
            t = p.text
            if t.strip() == "":
                compat.append({"level": "empty", "runs": [("", False)], "text": t})
            else:
                lvl = _classify_block(t, _ea_of(p), _has_numpr(p), _numid_of(p))
                compat.append({"level": lvl, "runs": [(t, False)], "text": t})

    return {"title": title, "unit": unit, "date": date,
            "dx": path, "doc": doc, "blocks": blocks, "paras": compat}


def clean_signoff(paras):
    """删除末尾的落款（单位 + 日期）。返回过滤后的列表。"""
    # 去掉尾部空段
    while paras and paras[-1]["text"].strip() == "":
        paras.pop()
    if not paras:
        return paras
    # 末段若为日期 → 删
    if _DATE_RE.search(paras[-1]["text"]):
        paras.pop()
        # 再删可能的空段
        while paras and paras[-1]["text"].strip() == "":
            paras.pop()
        # 上一段若为"单位"特征短句 → 删
        if paras:
            last = paras[-1]["text"].strip()
            if (len(last) <= 20 and any(h in last for h in _UNIT_HINT)) or \
               last.endswith(("政府", "局", "委", "公司", "办公室", "中心")):
                paras.pop()
    # 再清一次尾部空段
    while paras and paras[-1]["text"].strip() == "":
        paras.pop()
    return paras


# ============================================================
# 构建：封面 / 目录 / 正文
# ============================================================
def _setup_section(section):
    """仅设置页面尺寸与页边距。封面/目录节不设页脚（无页码）；
    正文节的页脚由 _add_number_footer 单独添加。
    同时设置 WPS 公文文档网格（type=lines, linePitch=312 ≈ 15.6pt），
    使表格单元格行距吸附到网格，与源文档（多数为 312）渲染一致；
    固定行距（28.5pt）段落不受网格影响。"""
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(MARGIN_LR_CM)
    section.right_margin = Cm(MARGIN_LR_CM)
    section.top_margin = Cm(MARGIN_TB_CM)
    section.bottom_margin = Cm(MARGIN_TB_CM)
    sectPr = section._sectPr
    dg = sectPr.find(qn("w:docGrid"))
    if dg is None:
        dg = OxmlElement("w:docGrid")
        _insert_sectpr_child(sectPr, dg)
    dg.set(qn("w:type"), "lines")
    dg.set(qn("w:linePitch"), "312")
    dg.set(qn("w:charSpace"), "0")


def _enable_even_odd(document):
    settings = document.settings.element
    if settings.find(qn("w:evenAndOddHeaders")) is None:
        settings.append(OxmlElement("w:evenAndOddHeaders"))


_SECTPR_ORDER = ["footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
                 "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign",
                 "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter",
                 "docGrid", "printerSettings", "sectPrChange"]


def _insert_sectpr_child(sectPr, el):
    """按 OOXML CT_SectPr 规范顺序插入子元素（docGrid 必须排在 pgNumType 之后，
    顺序违规会导致 WPS 忽略文档网格、表格行距不吸附）。"""
    tag = etree.QName(el).localname
    idx = _SECTPR_ORDER.index(tag) if tag in _SECTPR_ORDER else len(_SECTPR_ORDER)
    for child in sectPr:
        ct = etree.QName(child).localname
        if ct in _SECTPR_ORDER and _SECTPR_ORDER.index(ct) > idx:
            child.addprevious(el)
            return
    sectPr.append(el)


def _set_pg_num_start(section, start=1):
    sectPr = section._sectPr
    pgNum = sectPr.find(qn("w:pgNumType"))
    if pgNum is None:
        pgNum = OxmlElement("w:pgNumType")
        _insert_sectpr_child(sectPr, pgNum)
    pgNum.set(qn("w:start"), str(start))
    pgNum.set(qn("w:fmt"), "decimal")


def _add_number_footer(section, which="footer"):
    ftr = getattr(section, which)
    ftr.is_linked_to_previous = False
    for p in list(ftr.paragraphs):
        p._p.getparent().remove(p._p)
    p = ftr.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_exact_line(p, LINE_PT)
    _add_page_field(p)


def build_cover(document, cfg):
    """封面：左上角『会议资料 不要外传』+ 大标题两行（会议名称 / 汇报材料，方正小标宋简体）
    + 底部日期。

    用户在「汇编总标题」填会议名称（如 全市服务业大会），封面生成两段：
        全市服务业大会
        汇报材料
    若输入本身以「汇报材料」结尾（如 全市服务业大会汇报材料），则自动拆分为两行，不重复。
    """
    cover_title = (cfg.get("cover_title") or "汇编材料").strip()
    date_text = (cfg.get("date_text") or "").strip()
    title_size = cfg.get("cover_title_size", SIZE_TITLE)

    # 拆分：会议名称 + 汇报材料（输入已含结尾"汇报材料"时拆开，避免重复出现两次）
    if cover_title.endswith("汇报材料") and cover_title != "汇报材料":
        meeting = cover_title[:-len("汇报材料")].strip()
    else:
        meeting = cover_title
    title_lines = [line for line in (meeting, "汇报材料") if line]

    # 左上角：会议资料 / 不要外传（黑体、加粗、左对齐，依模板保留）
    _add_para(document, "会议资料", align=WD_ALIGN_PARAGRAPH.LEFT,
              eastasia="黑体", ascii_font="黑体", size=SIZE_BODY, bold=True)
    _add_para(document, "不要外传", align=WD_ALIGN_PARAGRAPH.LEFT,
              eastasia="黑体", ascii_font="黑体", size=SIZE_BODY, bold=True)
    # 顶部留白把大标题推到中部（留白过大会把日期挤到第二页，控制在安全范围）
    for _ in range(5):
        _add_para(document, "", align=WD_ALIGN_PARAGRAPH.CENTER)
    # 大标题：会议名称、汇报材料 各一段（同字体字号，居中）
    for line in title_lines:
        _add_para(document, line, align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=FONT_TITLE, ascii_font=FONT_TITLE, size=title_size, bold=True)
    # 底部留白把日期推到页脚上方（与封面同页）
    for _ in range(6):
        _add_para(document, "", align=WD_ALIGN_PARAGRAPH.CENTER)
    if date_text:
        _add_para(document, date_text, align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=FONT_TITLE, ascii_font=FONT_TITLE, size=SIZE_DATE)


def build_toc(document, materials):
    """目录：『目  录』标题 + WPS/Word 自动目录域（按大纲级别抓取材料标题并生成页码）。"""
    _add_toc_field(document)


def build_body(document, materials, options, ctx, fmt=None):
    """正文：每份材料 = 材料标题 + 单位行 + 内容（克隆源文档块元素）。

    关键设计（修复图片/表格/自动编号丢失、标题吞正文等）：
      - 内容采用『克隆源块元素』而非『按文本重建』：原文档的段落、表格、
        图片(drawing)、自动编号(numPr) 全部原样保留，再在 run/段落级套用
        汇编字体与间距。表格与图片保持原文档格式（不套汇编正文格式）。
      - 克隆出的图片媒体与自动编号定义经 ctx 登记后，随输出包补丁写入，
        使序号（一、/（一）等）与图片在汇编后正常显示。
    fmt：界面可调格式 {toc,body,h1,h2: {size,line_mode,line_value}}（已归一化），
    仅作用于 正文/一级/二级标题（及 body 空行）；表格/图片/h3 不受影响。
    """
    fmt = _merge_fmt(fmt)
    sizes = {k: fmt[k]["size"] for k in ("body", "h1", "h2")}
    from docx.text.paragraph import Paragraph
    out_body = document.element.body
    for i, m in enumerate(materials, 1):
        # 材料标题（方正小标宋简体、二号、居中、不加粗、段前段后间距 0；设大纲级别供自动目录抓取）
        title_p = _add_para_runs(document, [(m["title"], False)], "title",
                                align=WD_ALIGN_PARAGRAPH.CENTER,
                                space_before=Pt(0), space_after=Pt(0))
        _set_outline_level(title_p, 1)
        # 单位行（楷体_GB2312、三号、居中、段前段后间距 0）
        unit = (m.get("unit") or "").strip()
        if unit:
            _add_para_runs(document, [(unit, False)], "unit",
                           align=WD_ALIGN_PARAGRAPH.CENTER,
                           space_before=Pt(0), space_after=Pt(0))
        # 单位（或标题）与正文第一段之间空一行（需求：空一行正文；行距跟随正文设置）
        _add_para_runs(document, [("", False)], "body",
                       align=WD_ALIGN_PARAGRAPH.LEFT,
                       space_before=Pt(0), space_after=Pt(0),
                       line_mode=fmt["body"]["line_mode"],
                       line_pt=fmt["body"]["line_value"],
                       sizes=sizes)
        # 内容块：克隆源元素（保留图片/表格/编号），再套字体/间距
        blocks = m.get("blocks") or []
        src_dx = m.get("dx")
        for kind, el, p in blocks:
            clone = copy.deepcopy(el)
            _remap_block_refs(clone, src_dx, ctx)   # 图片 rId / 编号 numId / 样式 重映射
            if kind == "tbl":
                # 表格：原样保留格式；把单元格段落/run 的有效格式物化为直接格式，
                # 使其行距/段后间距/字体等与源文档一致（不随输出默认样式改变）。
                # 随后去除 snapToGrid/adjustRightInd，解除单元格与输出文档网格的耦合，
                # 避免行高被输出网格（与源不同）撑高导致与源文档不一致。
                _fix_cell_inheritance(clone, src_dx, ctx)
                _strip_table_grid_deps(clone)
            # 关键：必须插到 body 末尾 <w:sectPr> 之前，否则正文会落到正文节之外，
            # 导致"只有标题没有内容"、内容与标题错位
            out_body.insert_element_before(clone, "w:sectPr")
            if kind == "tbl":
                continue
            # 段落：套用层级字体（图片 run 跳过），保留 numPr/drawing
            para = Paragraph(clone, document)
            if _has_image(clone):
                # 含图片/嵌入对象的段落：保持源段落格式（对齐/行距/缩进等），不套汇编统一格式
                _strip_outline_level(para)   # 但源文档自带的大纲级别仍须剥离，防止混入目录
                continue
            lvl = _classify_block(p.text, _ea_of(p), _has_numpr(p), _numid_of(p))
            # 目录只收材料标题：克隆段落一律剥离大纲级别（含源文档自带的级别，
            # 如某些源正文段带 1 级大纲——更新目录域后会混入目录）
            _strip_outline_level(para)
            if lvl == "empty":
                _apply_para_uniform(para, indent=0,
                                    line_mode=fmt["body"]["line_mode"],
                                    line_value=fmt["body"]["line_value"],
                                    size_pt=fmt["body"]["size"])
                continue
            if lvl == "h3":
                # 三级标题（1. / （1）/ ① 等编号开头）：按用户要求保留原文格式，
                # 不套统一行距/缩进/对齐/字体
                continue
            # 行距（默认固定 28.5 磅，界面可调倍数/固定值）、段前段后 0、
            # 左右/悬挂缩进 0（表格单元格除外）
            if lvl in ("h1", "h2"):
                f = fmt[lvl]
                _apply_para_uniform(para, indent=2,
                                    line_mode=f["line_mode"], line_value=f["line_value"],
                                    size_pt=f["size"])
                # 一级/二级标题：左对齐（不再设大纲级别——旧版会使 h1/h2 进入目录）
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:  # body
                f = fmt["body"]
                indent = 0 if _is_salutation(p.text) else 2
                _apply_para_uniform(para, indent=indent,
                                    line_mode=f["line_mode"], line_value=f["line_value"],
                                    size_pt=f["size"])
                # 正文：两段对齐
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _apply_clone_fonts(para, lvl, sizes)
        # 材料之间分页（每份材料另起一页；最后一份后不再加分页，避免空白尾页）
        if options.get("each_material_new_page", True) and i < len(materials):
            document.add_page_break()
        else:
            _add_para(document, "", align=WD_ALIGN_PARAGRAPH.LEFT)


# ============================================================
# 资源登记与包补丁（图片媒体 + 自动编号定义）
# ============================================================
class AssemblyCtx:
    """跨材料收集需要并入输出包的外部资源：图片媒体 + 自动编号定义。"""

    def __init__(self):
        self.media = {}        # content_hash -> (new_rid, rel_target, blob, ext, ctype)
        self._media_ctr = 1000
        self._num = {}         # (src_dx, old_numid) -> new_numid
        self._abnum = {}       # (src_dx, old_abnumid) -> new_abnumid
        self._num_ctr = 1000
        self._src_cache = {}   # src_dx -> (abnum_by_id, num_by_id)
        self.num_needed = []   # (src_dx, old_numid, new_numid, new_abnumid, src_abnumid)
        self._style = {}       # (src_dx, old_styleid) -> new_styleid
        self._style_ctr = 0
        self._src_style_cache = {}   # src_dx -> {"docDefaults": el|None, "styles": {id: el}}
        self._styles_needed = []     # (src_dx, old_id, new_id)
        self._first_src = None       # 首个被处理源（用于 docDefaults/Normal 默认替换）

    def _src_styles(self, src_dx):
        if src_dx in self._src_style_cache:
            return self._src_style_cache[src_dx]
        info = {"docDefaults": None, "styles": {}}
        try:
            with zipfile.ZipFile(src_dx, "r") as z:
                if "word/styles.xml" in z.namelist():
                    root = etree.fromstring(z.read("word/styles.xml"))
                    for el in root:
                        tag = etree.QName(el).localname
                        if tag == "docDefaults":
                            info["docDefaults"] = el
                        elif tag == "style":
                            sid = el.get(qn("w:styleId"))
                            if sid:
                                info["styles"][sid] = el
        except Exception:
            pass
        self._src_style_cache[src_dx] = info
        return info

    def register_style(self, src_dx, old_id):
        """登记样式（pStyle/rStyle/tblStyle）并重映射 styleId，返回新 id。

        同时把基于该样式的 basedOn/link/next 链上的样式也一并登记，
        并在并入输出 styles.xml 时统一改写内部引用。
        """
        if not old_id:
            return None
        if self._first_src is None:
            self._first_src = src_dx
        key = (src_dx, old_id)
        if key in self._style:
            return self._style[key]
        info = self._src_styles(src_dx)
        if old_id not in info["styles"]:
            self._style[key] = None
            return None
        self._style_ctr += 1
        new_id = "s%d_%s" % (self._style_ctr, old_id)
        self._style[key] = new_id
        self._styles_needed.append((src_dx, old_id, new_id))
        # 传递登记引用链
        for child in info["styles"][old_id]:
            ct = etree.QName(child).localname
            if ct in ("basedOn", "link", "next"):
                ref = child.get(qn("w:val"))
                if ref and ref != old_id:
                    self.register_style(src_dx, ref)
        return new_id

    def _src_numbering(self, src_dx):
        if src_dx in self._src_cache:
            return self._src_cache[src_dx]
        ab, nm = {}, {}
        try:
            with zipfile.ZipFile(src_dx, "r") as z:
                if "word/numbering.xml" in z.namelist():
                    root = etree.fromstring(z.read("word/numbering.xml"))
                    for el in root:
                        tag = etree.QName(el).localname
                        if tag == "abstractNum":
                            ab[el.get(qn("w:abstractNumId"))] = el
                        elif tag == "num":
                            nm[el.get(qn("w:numId"))] = el
        except Exception:
            pass
        self._src_cache[src_dx] = (ab, nm)
        return ab, nm

    def register_relationship(self, src_dx, old_rid):
        """登记任意外部关系（图片 media / OLE embeddings），返回新 rId（同内容去重）。

        图片可能以多种方式引用：DML(a:blip r:embed/link)、VML(v:imagedata r:id)、
        OLE(o:OLEObject r:id)。统一按关系目标复制二进制并写入输出包。
        """
        try:
            with zipfile.ZipFile(src_dx, "r") as z:
                rel_root = etree.fromstring(z.read("word/_rels/document.xml.rels"))
                target, rtype = None, None
                for rel in rel_root:
                    # .rels 的属性（Id/Type/Target）属于无命名空间
                    if rel.get("Id") == old_rid:
                        target = rel.get("Target")
                        rtype = rel.get("Type")
                        break
                if not target:
                    return None
                t = target.replace("\\", "/")
                # 源内读取路径：相对 word/（Target 默认相对 word/）或包根（../、/ 开头）
                if t.startswith("/"):
                    src_read = t.lstrip("/")
                elif t.startswith(".."):
                    src_read = "/".join(p for p in t.split("/") if p not in ("..", "."))
                else:
                    src_read = "word/" + t
                data = z.read(src_read)
                name = src_read.rsplit("/", 1)[-1]
                ext = name.rsplit(".", 1)[-1].lower()
        except Exception:
            return None
        h = hashlib.md5(data).hexdigest()
        if h in self.media:
            return self.media[h][0]
        is_ole = (rtype or "").lower().endswith("oleobject") or src_read.lower().startswith("word/embeddings/")
        is_media = src_read.lower().startswith("word/media/") or ext in (
            "png", "jpg", "jpeg", "gif", "bmp", "emf", "wmf", "tif", "tiff", "svg")
        self._media_ctr += 1
        new_rid = "rId%d" % self._media_ctr
        if is_ole:
            out_rel = "embeddings/asm_%d.bin" % self._media_ctr
            ctype = None
        else:
            out_rel = "media/asm_%d.%s" % (self._media_ctr, ext if ext else "png")
            ctype = {
                "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                "gif": "image/gif", "bmp": "image/bmp", "emf": "image/x-emf",
                "wmf": "image/x-wmf", "tif": "image/tiff", "tiff": "image/tiff",
                "svg": "image/svg+xml",
            }.get(ext, "image/png")
        self.media[h] = (new_rid, out_rel, data, ext, rtype, is_media)
        return new_rid

    @staticmethod
    def media_ctype(ext):
        return {
            "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
            "gif": "image/gif", "bmp": "image/bmp", "emf": "image/x-emf",
            "wmf": "image/x-wmf", "tif": "image/tiff", "tiff": "image/tiff",
            "svg": "image/svg+xml",
        }.get((ext or "").lower(), "image/png")

    def register_numid(self, src_dx, old_numid):
        """登记自动编号 numId，返回新 numId（同时登记其 abstractNum）。"""
        key = (src_dx, old_numid)
        if key in self._num:
            return self._num[key]
        ab, nm = self._src_numbering(src_dx)
        num_el = nm.get(old_numid)
        if num_el is None:
            self._num[key] = None
            return None
        src_ab = None
        for c in num_el:
            if etree.QName(c).localname == "abstractNumId":
                src_ab = c.get(qn("w:val"))
                break
        if src_ab is None:
            self._num[key] = None
            return None
        if (src_dx, src_ab) not in self._abnum:
            self._num_ctr += 1
            self._abnum[(src_dx, src_ab)] = self._num_ctr
        new_ab = self._abnum[(src_dx, src_ab)]
        self._num_ctr += 1
        new_num = self._num_ctr
        self._num[key] = new_num
        self.num_needed.append((src_dx, old_numid, new_num, new_ab, src_ab))
        return new_num


def _remap_block_refs(clone, src_dx, ctx):
    """重映射克隆块内的外部引用：图片（DML a:blip / VML v:imagedata / OLE o:OLEObject）
    与自动编号 numId，使其指向输出包中的新资源。无法解析的 numId 移除 numPr，
    避免悬空引用导致序号错乱。"""
    if src_dx is None:
        return
    # 收集图片/对象引用（先收集再改，避免迭代中删除）
    refs = []
    for blip in clone.iter("{%s}blip" % A_NS):
        for attr in ("embed", "link"):
            rid = blip.get("{%s}%s" % (R_NS, attr))
            if rid:
                refs.append((blip, "{%s}%s" % (R_NS, attr), rid))
    for img in clone.iter("{%s}imagedata" % V_NS):
        rid = img.get("{%s}id" % R_NS)
        if rid:
            refs.append((img, "{%s}id" % R_NS, rid))
    for ole in clone.iter("{%s}OLEObject" % O_NS):
        rid = ole.get("{%s}id" % R_NS)
        if rid:
            refs.append((ole, "{%s}id" % R_NS, rid))
    for el, attr, rid in refs:
        nrid = ctx.register_relationship(src_dx, rid)
        if nrid:
            el.set(attr, nrid)
    # 自动编号：w:numPr/w:numId/@w:val
    for np in list(clone.iter(qn("w:numPr"))):
        nid = np.find(qn("w:numId"))
        if nid is None:
            continue
        old = nid.get(qn("w:val"))
        if old is None:
            continue
        nnew = ctx.register_numid(src_dx, old)
        if nnew is not None:
            nid.set(qn("w:val"), str(nnew))
        else:
            # 源中无该 numId 定义（如 WPS 转换丢失）：移除 numPr 防悬空
            pPr = np.getparent()
            if pPr is not None:
                pPr.remove(np)
    # 样式引用：w:pStyle / w:rStyle（含表格单元格内段落）与 w:tblStyle
    #
    # 与上面 numId 同理：源文档可能引用自身 styles.xml 里**并不存在**的样式
    # （WPS 转换常见，例如段落引用 s1_Normal，而源 styles.xml 从未定义它）。
    # 这种悬空引用 Word / WPS 静默容忍，但 LibreOffice（本程序预览与 PDF
    # 导出的引擎）会据此丢弃该段落的**直接 w:ind**——首行缩进/左缩进整段失效，
    # 公文版式直接走样。解析不到就删除该引用，段落回落到默认样式，
    # 与 Word 下的视觉效果一致。
    for tag in ("pStyle", "rStyle"):
        for el in list(clone.iter(qn("w:" + tag))):
            v = el.get(qn("w:val"))
            if not v:
                continue
            nv = ctx.register_style(src_dx, v)
            if nv:
                el.set(qn("w:val"), nv)
            else:
                par = el.getparent()
                if par is not None:
                    par.remove(el)
    for tp in list(clone.iter(qn("w:tblPr"))):
        ts = tp.find(qn("w:tblStyle"))
        if ts is None:
            continue
        v = ts.get(qn("w:val"))
        if not v:
            continue
        nv = ctx.register_style(src_dx, v)
        if nv:
            ts.set(qn("w:val"), nv)
        else:
            tp.remove(ts)


def _cell_style_chain(info, p_el):
    """源文档段落样式链：从最具体到最不具体（pStyle→basedOn→…→Normal→docDefaults）。"""
    chain = []
    pPr = p_el.find(qn("w:pPr"))
    cur = None
    if pPr is not None:
        ps = pPr.find(qn("w:pStyle"))
        cur = ps.get(qn("w:val")) if ps is not None else None
    seen = set()
    while cur and cur not in seen:
        seen.add(cur)
        st = info["styles"].get(cur)
        if st is None:
            break
        chain.append(st)
        nxt = None
        for c in st:
            if etree.QName(c).localname == "basedOn":
                nxt = c.get(qn("w:val"))
        cur = nxt
    if "Normal" in info["styles"]:
        chain.append(info["styles"]["Normal"])
    if info["docDefaults"] is not None:
        chain.append(info["docDefaults"])
    return chain


def _materialize_missing(pPr, qtag, chain):
    """若 pPr 缺少 qtag 子元素，则沿链（最具体在前）找第一个定义复制进来。
    链上也没有定义时保持缺省（不写入默认值），让应用按无定义渲染——
    与源文档一致（例如表格单元格行距：WPS 对无 w:spacing 的单元格使用
    自己的表格默认行距，显式写入单倍行距反而会改变渲染）。"""
    if pPr.find(qn(qtag)) is not None:
        return
    for st in chain:
        ppr = st.find(qn("w:pPr"))
        if ppr is None:
            continue
        src = ppr.find(qn(qtag))
        if src is not None:
            _insert_ppr_child(pPr, copy.deepcopy(src))
            return


def _materialize_run_fonts(run_el, chain):
    """给缺少字体/字号/加粗/斜体的 run 物化源链上的定义，使单元格 run 与源一致。"""
    rPr = run_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run_el.insert(0, rPr)
    for qtag in ("w:rFonts", "w:sz", "w:szCs", "w:b", "w:i"):
        if rPr.find(qn(qtag)) is not None:
            continue
        for st in chain:
            rr = st.find(qn("w:rPr"))
            if rr is None:
                continue
            src = rr.find(qn(qtag))
            if src is not None:
                # 按 CT_RPr 规范顺序插入，避免 Word 判定文档损坏
                _insert_rpr_child(rPr, copy.deepcopy(src))
                break


def _fix_cell_inheritance(tbl_el, src_dx, ctx):
    """表格克隆后，把源文档单元格段落/run 的『有效格式』物化为直接格式：
    行距/段前段后/缩进/对齐/字体/字号 等不再依赖输出文档的默认样式，
    保证表格与源文档完全一致。"""
    if not src_dx:
        return
    info = ctx._src_styles(src_dx)
    for tc in tbl_el.iter(qn("w:tc")):
        for p in tc.iter(qn("w:p")):
            pPr = p.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                p.insert(0, pPr)
            chain = _cell_style_chain(info, p)
            # 仅当源链上确有定义时才物化；无定义则不写（保留 WPS 表格默认行距渲染）
            _materialize_missing(pPr, "w:spacing", chain)
            _materialize_missing(pPr, "w:ind", chain)
            _materialize_missing(pPr, "w:jc", chain)
            for r in p.findall(qn("w:r")):
                _materialize_run_fonts(r, chain)


def _has_image(el):
    """判断块元素是否含图片/嵌入对象（DML w:drawing、VML v:imagedata/w:pict、OLE o:OLEObject）。"""
    return (el.find(".//" + qn("w:drawing")) is not None
            or el.find(".//{%s}imagedata" % V_NS) is not None
            or el.find(".//{%s}OLEObject" % O_NS) is not None
            or el.find(".//" + qn("w:pict")) is not None
            or el.find(".//" + qn("w:object")) is not None)


def _strip_table_grid_deps(tbl_el):
    """去除表格单元格段落的 `w:snapToGrid`（及 `w:adjustRightInd`）标志。

    源文档的表格单元格段落常带 `snapToGrid` 用于贴合源文档自身的网格；
    源文档与我们输出文档的 `docGrid` linePitch 不同时，WPS 会按输出网格重新吸附，
    导致原本 `line=570 exact`（28.5 磅）的行高被网格撑高，表格行间距/换行与源不一致。
    去除该标志后，行高完全由段落自身的 `line` 决定，与文档网格解耦。
    """
    for p in tbl_el.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            continue
        for tag in ("snapToGrid", "adjustRightInd"):
            el = pPr.find(qn("w:" + tag))
            if el is not None:
                pPr.remove(el)


def _clone_run_fonts(run, level, sizes=None):
    """对单个文本 run 套用层级字体；图片/OLE run 与无文本 run 跳过。"""
    rEl = run._element
    if rEl.find(qn("w:drawing")) is not None or rEl.find(qn("w:object")) is not None:
        return
    if rEl.find(qn("w:t")) is None and rEl.find(qn("w:tab")) is None:
        return
    ea, ascii_font, size, bold_default = _level_fonts(level, sizes)
    _set_run_fonts(run, eastasia=ea, ascii_font=ascii_font, hansi=ascii_font,
                   size=size, bold=(bold_default if bold_default is not None else None))


_PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
              "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
              "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
              "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
              "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
              "suppressOverlap", "jc", "textDirection", "textAlignment",
              "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
              "sectPr", "pPrChange"]


def _insert_ppr_child(pPr, el):
    """按 OOXML CT_PPr 规范顺序插入子元素（w:ind 须排在 w:spacing 之后、w:jc 之前）。"""
    tag = etree.QName(el).localname
    idx = _PPR_ORDER.index(tag) if tag in _PPR_ORDER else len(_PPR_ORDER)
    for child in pPr:
        ct = etree.QName(child).localname
        if ct in _PPR_ORDER and _PPR_ORDER.index(ct) > idx:
            child.addprevious(el)
            return
    pPr.append(el)


def _zero_horizontal_indents(para):
    """文本之前缩进 / 文本之后缩进 / 悬挂缩进 显式置 0。

    要点：必须「显式置 0」而非删除属性。源段落正是靠 w:left="0" w:leftChars="0"
    抵消段落样式（WPS 的 toc 2 / toc 4 等样式自带 left/leftChars 缩进，实测有
    2 字符、2.38 厘米、6 字符等）与编号定义里的缩进；删除这些属性会让样式/编号的
    缩进重新生效，表现为「整段缩进 N 个字符」。

    但 w:hanging / w:hangingChars 必须**删除**而不是显式置 0（见 _set_first_line_indent）：
      - WPS：w:hangingChars 只要存在（哪怕 0）就会让 firstLineChars 失效 → 首行缩进整体丢失；
      - LibreOffice：w:hanging 只要存在（哪怕 0）就会让 firstLine 失效 → 预览里首行不缩进。
    悬挂缩进因此统一由 _set_first_line_indent 显式写 firstLine 来抵消样式/编号里的悬挂值
    （样式链上的 hangingChars 已在 _patch_package 并入样式/编号时统一剥离）。
    """
    pPr = para._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        _insert_ppr_child(pPr, ind)
    for a in ("w:left", "w:leftChars", "w:right", "w:rightChars"):
        ind.set(qn(a), "0")
    for a in ("w:hanging", "w:hangingChars"):
        if ind.get(qn(a)) is not None:
            del ind.attrib[qn(a)]


def _apply_para_uniform(para, indent=2, line_mode="exact", line_value=LINE_PT, size_pt=None):
    """统一正文级段落：行距（默认固定 28.5 磅，可调倍数/固定值）、段前段后 0、
    首行缩进 indent 字符、左右/悬挂缩进 0。表格单元格与图片段不走此函数。

    size_pt 为该段落实际使用的字号（磅），用于把「N 字符缩进」折算成等值 twips
    （LibreOffice 只认 twips，见 _set_first_line_indent）。
    """
    _set_line(para, line_mode, line_value)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    _set_first_line_indent(para, indent, size_pt)
    _zero_horizontal_indents(para)


def _apply_clone_fonts(para, lvl, sizes=None):
    """对克隆段落按层级套字体。

    一级标题：整段黑体不加粗。
    二级标题：从（一）标识到该句第一个句号（含句号）为楷体_GB2312 加粗（标题字），
              句号之后的文字为正文（仿宋_GB2312 中文 / 宋体 西文，保留原加粗）。
    正文：仿宋/宋体，保留原加粗。
    sizes 为界面可调字号（{body,h1,h2: 磅}）。
    """
    runs = para.runs
    if lvl != "h2":
        for run in runs:
            _clone_run_fonts(run, lvl, sizes)
        return
    full = "".join(r.text or "" for r in runs)
    cut = full.find("。")
    if cut == -1:
        for run in runs:
            _clone_run_fonts(run, "h2", sizes)
        return
    cut_end = cut + 1
    from docx.text.run import Run
    pos = 0
    for run in list(runs):
        t = run.text or ""
        if not t:
            continue
        start, end = pos, pos + len(t)
        pos = end
        if end <= cut_end:
            _clone_run_fonts(run, "h2", sizes)
        elif start >= cut_end:
            _clone_run_fonts(run, "body", sizes)
        else:
            # 一个 run 跨句号：拆成标题部分(当前 run) + 正文部分(新 run)
            s = cut_end - start
            head_t, tail_t = t[:s], t[s:]
            new_r = copy.deepcopy(run._element)   # 先复制（含原完整文本）
            run.text = head_t                     # 当前 run 只保留标题部分
            ts = new_r.findall(qn("w:t"))
            if ts:
                for i, nt in enumerate(ts):
                    nt.text = tail_t if i == 0 else ""
            else:
                t_el = OxmlElement("w:t")
                t_el.set(qn("xml:space"), "preserve")
                t_el.text = tail_t
                new_r.append(t_el)
            run._element.addnext(new_r)
            _clone_run_fonts(run, "h2", sizes)
            _clone_run_fonts(Run(new_r, para), "body", sizes)


def _neutralize_default_styles(st_root):
    """输出默认样式中性化：清空 docDefaults 的 pPrDefault 与 Normal 的
    行距/缩进/对齐，使无直接格式的内容按『无定义』渲染（与各源一致）；
    并把默认字体置为中文常用（宋体 / Times New Roman）。"""
    dd = st_root.find(qn("w:docDefaults"))
    if dd is not None:
        ppr = dd.find(qn("w:pPrDefault") + "/" + qn("w:pPr"))
        if ppr is not None:
            for tag in ("w:spacing", "w:ind", "w:jc"):
                el = ppr.find(qn(tag))
                if el is not None:
                    ppr.remove(el)
        rpr = dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr"))
        if rpr is not None:
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None:
                rf.set(qn("w:ascii"), "Times New Roman")
                rf.set(qn("w:hAnsi"), "Times New Roman")
                rf.set(qn("w:eastAsia"), "宋体")
                rf.set(qn("w:cs"), "Times New Roman")
    for st in st_root:
        if etree.QName(st).localname != "style":
            continue
        if st.get(qn("w:styleId")) == "Normal":
            ppr = st.find(qn("w:pPr"))
            if ppr is not None:
                for tag in ("w:spacing", "w:ind", "w:jc"):
                    el = ppr.find(qn(tag))
                    if el is not None:
                        ppr.remove(el)
            rpr = st.find(qn("w:rPr"))
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None:
                    rf.set(qn("w:ascii"), "Times New Roman")
                    rf.set(qn("w:hAnsi"), "Times New Roman")
                    rf.set(qn("w:eastAsia"), "宋体")
            break


def _patch_package(output_path, ctx):
    """保存后补丁：写入引用图片媒体（更新 rels/Content_Types），
    并把所需编号定义(abstractNum/num)映射后并入输出 numbering.xml，
    使原文档的自动编号序号（一、/（一）等）在汇编后重新显示。"""
    if not ctx.media and not ctx.num_needed and not ctx._styles_needed and ctx._first_src is None:
        return
    with zipfile.ZipFile(output_path, "r") as z:
        parts = {n: z.read(n) for n in z.namelist()}

    # ---- 媒体 / OLE ----
    if ctx.media:
        rel_root = etree.fromstring(parts["word/_rels/document.xml.rels"])
        ct_root = etree.fromstring(parts["[Content_Types].xml"])
        for (new_rid, out_rel, data, ext, rtype, is_media) in ctx.media.values():
            parts["word/" + out_rel] = data
            rel = etree.SubElement(rel_root, "{%s}Relationship" % PKG_R_NS)
            rel.set("Id", new_rid)
            rel.set("Type", rtype or "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            rel.set("Target", out_rel)
            # OPC 要求包内每个部件都有内容类型（Default 按扩展名 或 Override 按部件名）。
            # 媒体与 OLE 嵌入对象都必须声明，缺失会让 Word 判定"文件已损坏"并拒绝打开。
            # 注意：须按输出部件的实际扩展名声明——OLE 嵌入对象源可能是 .xls/.docx 等，
            # 但统一落盘为 embeddings/asm_N.bin，按源扩展名声明会漏掉 bin。
            out_ext = out_rel.rsplit(".", 1)[-1].lower() if "." in out_rel else ext
            if out_ext:
                if not any(etree.QName(d).localname == "Default" and d.get("Extension", "").lower() == out_ext
                           for d in ct_root):
                    d = etree.SubElement(ct_root, "{%s}Default" % CT_NS)
                    d.set("Extension", out_ext)
                    d.set("ContentType",
                          ctx.media_ctype(ext) if is_media
                          else "application/vnd.openxmlformats-officedocument.oleObject")
        parts["word/_rels/document.xml.rels"] = etree.tostring(rel_root, xml_declaration=True, encoding="UTF-8")
        parts["[Content_Types].xml"] = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8")

    # ---- 编号 ----
    if ctx.num_needed:
        if "word/numbering.xml" in parts:
            num_root = etree.fromstring(parts["word/numbering.xml"])
            created = False
        else:
            num_root = etree.fromstring(
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>')
            created = True
        W = "{%s}" % W_NS
        # 复制所需 abstractNum（按 (src_dx, src_ab) 去重）
        copied_ab = set()
        for (src_dx, old_numid, new_num, new_ab, src_ab) in ctx.num_needed:
            if (src_dx, src_ab) in copied_ab:
                continue
            s_ab, _ = ctx._src_numbering(src_dx)
            src_el = s_ab.get(src_ab)
            if src_el is not None:
                new_el = copy.deepcopy(src_el)
                new_el.set(qn("w:abstractNumId"), str(new_ab))
                # 同样式：剥离编号层级缩进里的 hangingChars（会使 firstLineChars 失效）
                for ind_el in new_el.iter(qn("w:ind")):
                    if ind_el.get(qn("w:hangingChars")) is not None:
                        del ind_el.attrib[qn("w:hangingChars")]
                num_root.append(new_el)
            copied_ab.add((src_dx, src_ab))
        # 复制所需 num（重映射 numId 与 abstractNumId 引用）
        for (src_dx, old_numid, new_num, new_ab, src_ab) in ctx.num_needed:
            s_ab, s_nm = ctx._src_numbering(src_dx)
            src_el = s_nm.get(old_numid)
            if src_el is None:
                continue
            new_el = copy.deepcopy(src_el)
            new_el.set(qn("w:numId"), str(new_num))
            for c in new_el:
                if etree.QName(c).localname == "abstractNumId":
                    c.set(qn("w:val"), str(new_ab))
            num_root.append(new_el)
        parts["word/numbering.xml"] = etree.tostring(num_root, xml_declaration=True, encoding="UTF-8")
        if created:
            # 新建 numbering 还需补 Content-Types Override 与 relationship
            ct_root = etree.fromstring(parts["[Content_Types].xml"])
            if not any(etree.QName(d).localname == "Override" and d.get("PartName") == "/word/numbering.xml"
                       for d in ct_root):
                ov = etree.SubElement(ct_root, "{%s}Override" % CT_NS)
                ov.set("PartName", "/word/numbering.xml")
                ov.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
            parts["[Content_Types].xml"] = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8")
            rel_root = etree.fromstring(parts["word/_rels/document.xml.rels"])
            rid_max = 0
            for rel in rel_root:
                m = re.search(r"rId(\d+)", rel.get("Id", ""))
                if m:
                    rid_max = max(rid_max, int(m.group(1)))
            rel = etree.SubElement(rel_root, "{%s}Relationship" % PKG_R_NS)
            rel.set("Id", "rId%d" % (rid_max + 1))
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering")
            rel.set("Target", "numbering.xml")
            parts["word/_rels/document.xml.rels"] = etree.tostring(rel_root, xml_declaration=True, encoding="UTF-8")

    # ---- 样式：并入源文档样式（重映射 styleId，修正 basedOn/link/next），
    #      使克隆元素的 pStyle/rStyle/tblStyle 引用可解析 ----
    if ctx._styles_needed or ctx._first_src:
        styles_path = "word/styles.xml"
        if styles_path in parts:
            st_root = etree.fromstring(parts[styles_path])
        else:
            st_root = etree.fromstring(
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>')
        # 默认样式中性化：把输出 docDefaults 的 pPr 与 Normal 的 行距/缩进/对齐 清空
        # （与各源一致——源 docDefaults 多为空 pPrDefault），并置中文默认字体，
        # 使表格单元格等继承内容按"无定义"渲染，与源文档一致
        _neutralize_default_styles(st_root)
        # 收集输出原有的"默认样式"（按 type+name 匹配）：并入的源样式若与输出
        # 默认样式同名同 type，会被 WPS 视为第二个默认样式参与解析，导致
        # 表格 autofit 按被并入源的 Normal 字体度量计算列宽，单元格收窄、
        # 文字断行（根因：组合汇编下工贸表列宽收缩）。统一跳过，不并入，
        # 输出仍以原有的 Normal/TableNormal 等为准。
        existing_defaults = set()
        native_defaults = {}   # (type, name) -> 输出原生默认样式的 styleId
        for s in st_root:
            if etree.QName(s).localname != "style":
                continue
            stype = s.get(qn("w:type"))
            if s.get(qn("w:default")) != "1" or not stype:
                continue
            for c in s:
                if etree.QName(c).localname == "name":
                    nv = c.get(qn("w:val"))
                    if nv:
                        existing_defaults.add((stype, nv))
                        _sid = s.get(qn("w:styleId"))
                        if _sid:
                            native_defaults[(stype, nv)] = _sid
                    break
        # 追加所有需要并入的样式（新 styleId），并修正内部基于/链接/后续引用
        old2new = {(sdx, oid): nid for (sdx, oid, nid) in ctx._styles_needed}
        # 输出 styles.xml 里**原生已存在**的 styleId：源样式的 basedOn/link/next
        # 若指向它们（如 Normal / DefaultParagraphFont），引用可以直接保留。
        native_ids = set()
        for _s in st_root:
            if etree.QName(_s).localname == "style":
                _sid = _s.get(qn("w:styleId"))
                if _sid:
                    native_ids.add(_sid)

        def _would_skip(_src_el):
            """该源样式是否会因「与输出默认样式同名同 type」而被跳过不并入。"""
            _t = _src_el.get(qn("w:type"))
            _n = None
            for _c in _src_el:
                if etree.QName(_c).localname == "name":
                    _n = _c.get(qn("w:val"))
                    break
            return bool(_t and _n and (_t, _n) in existing_defaults)

        # 预扫描：① 真正会被并入的 styleId 集合；② 「被跳过样式的重定向表」。
        #
        # 关键：_remap_block_refs 早已把段落的 w:pStyle 改写成新 styleId，
        # 而被跳过的样式（源 Normal 与输出默认样式同名同 type）最终并不会写进
        # styles.xml —— 于是正文里就留下一批**悬空引用**。Word/WPS 静默容忍，
        # LibreOffice（本程序预览/PDF 导出引擎）则会据此丢弃段落的直接 w:ind，
        # 导致首行缩进整段失效。这类样式与输出原生默认样式语义等价，
        # 因此把引用重定向过去，既不留悬空引用，也不改变版式语义。
        emitted_ids = set()
        redirect = {}          # 被跳过样式的新 id -> 输出原生默认样式 id
        for (src_dx, old_id, new_id) in ctx._styles_needed:
            _se = ctx._src_styles(src_dx)["styles"].get(old_id)
            if _se is None:
                continue
            _t = _se.get(qn("w:type"))
            _n = None
            for _c in _se:
                if etree.QName(_c).localname == "name":
                    _n = _c.get(qn("w:val"))
                    break
            _k = (_t, _n) if (_t and _n) else None
            if _k and _k in existing_defaults:
                redirect[new_id] = native_defaults.get(_k, "Normal")
            else:
                emitted_ids.add(new_id)

        for (src_dx, old_id, new_id) in ctx._styles_needed:
            src_el = ctx._src_styles(src_dx)["styles"].get(old_id)
            if src_el is None:
                continue
            new_el = copy.deepcopy(src_el)
            new_el.set(qn("w:styleId"), new_id)
            # 剥离源样式的 default="1" 标志：并入的源样式与输出原有的 Normal /
            # Normal Table 等默认样式同名重复时，WPS 会按"最后一个 default"解析，
            # 导致组合汇编下表格 autofit 按被并入源的 Normal 字体度量计算列宽，
            # 单元格收窄、文字断行。统一把并入源样式的 default 属性去掉，
            # 保留其 rFonts/pPr/rPr 等格式定义，但默认角色仍归输出原有的 Normal。
            if new_el.get(qn("w:default")) is not None:
                del new_el.attrib[qn("w:default")]
            # 若该源样式与输出某个默认样式同名同 type（即便已剥 default 也会被
            # WPS 当作"非默认但同名 Normal"参与解析，引发列宽收缩），整条跳过。
            # 这类样式在源文档里就是 default 角色，被并入后已无意义。
            _src_type = new_el.get(qn("w:type"))
            _src_name = None
            for _c in new_el:
                if etree.QName(_c).localname == "name":
                    _src_name = _c.get(qn("w:val"))
                    break
            if _src_type and _src_name and (_src_type, _src_name) in existing_defaults:
                continue
            # 样式链上的 hangingChars 会让 WPS 忽略段落直接格式的 firstLineChars
            # （首行缩进失效），统一剥离（W(_zero_horizontal_indents) 已删直接属性）
            for ind_el in new_el.iter(qn("w:ind")):
                if ind_el.get(qn("w:hangingChars")) is not None:
                    del ind_el.attrib[qn("w:hangingChars")]
            # 引用改写：能解析到的重映射；指向输出原生样式（Normal 等）的保留；
            # 其余（源未定义 / 指向被跳过的样式）一律删除，避免悬空引用让
            # LibreOffice 丢弃段落直接 w:ind（首行缩进失效）。
            _drop = []
            for child in new_el:
                ct = etree.QName(child).localname
                if ct in ("basedOn", "link", "next"):
                    ref = child.get(qn("w:val"))
                    if not ref:
                        continue
                    if (src_dx, ref) in old2new:
                        _nid = old2new[(src_dx, ref)]
                        if _nid in emitted_ids:
                            child.set(qn("w:val"), _nid)
                        elif _nid in redirect:
                            child.set(qn("w:val"), redirect[_nid])
                        else:
                            _drop.append(child)
                    elif ref not in native_ids:
                        _drop.append(child)
                elif ct == "tblPr":
                    # 表格样式内部还可能引用其它表格样式（w:tblStyle），同样重映射
                    ts = child.find(qn("w:tblStyle"))
                    if ts is not None:
                        ref = ts.get(qn("w:val"))
                        if ref:
                            if (src_dx, ref) in old2new:
                                _nid = old2new[(src_dx, ref)]
                                if _nid in emitted_ids:
                                    ts.set(qn("w:val"), _nid)
                                elif _nid in redirect:
                                    ts.set(qn("w:val"), redirect[_nid])
                                else:
                                    _drop.append(ts)
                            elif ref not in native_ids:
                                _drop.append(ts)
            for _el in _drop:
                _par = _el.getparent()
                if _par is not None:
                    _par.remove(_el)
            st_root.append(new_el)
        # 最后一道保险：把正文 / 页眉页脚里指向「已被跳过样式」的引用，统一
        # 重定向到输出原生默认样式（源 Normal → 输出 Normal）。确保整个包中
        # 不存在任何解析不到的样式引用——LibreOffice 对悬空 pStyle 的报复
        # 就是丢弃该段落的直接 w:ind（首行缩进失效）。
        if redirect:
            for _pn in list(parts):
                if _pn != "word/document.xml" and not (
                        (_pn.startswith("word/header") or _pn.startswith("word/footer"))
                        and _pn.endswith(".xml")):
                    continue
                try:
                    _rt = etree.fromstring(parts[_pn])
                except Exception:
                    continue
                _hit = 0
                for _tag in ("pStyle", "rStyle", "tblStyle"):
                    for _el in _rt.iter(qn("w:" + _tag)):
                        _v = _el.get(qn("w:val"))
                        if _v in redirect:
                            _el.set(qn("w:val"), redirect[_v])
                            _hit += 1
                if _hit:
                    parts[_pn] = etree.tostring(_rt, xml_declaration=True, encoding="UTF-8")
        parts[styles_path] = etree.tostring(st_root, xml_declaration=True, encoding="UTF-8")

    # ---- settings.xml：采用首个源的完整设置（含 WPS 文档网格/行距兼容标志，
    #      使表格内行距吸附文档网格，与源文档渲染一致；仅加一个 adjustLineHeightInTable
    #      不够，需要整套 compat 标志），并确保保留 evenAndOddHeaders（双页翻面页码）----
    if ctx._first_src:
        try:
            with zipfile.ZipFile(ctx._first_src) as z:
                raw = z.read("word/settings.xml")
            sx = etree.fromstring(raw)
            if sx.find(qn("w:evenAndOddHeaders")) is None:
                sx.append(OxmlElement("w:evenAndOddHeaders"))
            parts["word/settings.xml"] = etree.tostring(sx, xml_declaration=True, encoding="UTF-8")
        except Exception:
            pass

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)


# ============================================================
# 主入口
# ============================================================
# 界面可调格式的默认值：目录内容/正文/一级/二级标题 全部三号 + 固定行距 28.5 磅。
# fmt 结构：{toc|body|h1|h2: {size(磅), line_mode('exact'|'multiple'), line_value}}
DEFAULT_FMT = {
    "toc":  {"size": 16.0, "line_mode": "exact", "line_value": 28.5},
    "body": {"size": 16.0, "line_mode": "exact", "line_value": 28.5},
    "h1":   {"size": 16.0, "line_mode": "exact", "line_value": 28.5},
    "h2":   {"size": 16.0, "line_mode": "exact", "line_value": 28.5},
}


def _merge_fmt(fmt):
    """把界面传入的 fmt 与默认值合并，产出四键齐全的归一化格式。

    键：toc / body / h1 / h2；每键 {size: 磅(float), line_mode: exact|multiple,
    line_value: float}。未提供或非法的字段回退默认（三号 + 固定 28.5）。
    """
    out = {}
    for k, d in DEFAULT_FMT.items():
        item = dict(d)
        u = (fmt or {}).get(k) or {}
        if u.get("size") is not None:
            try:
                item["size"] = float(u["size"])
            except (TypeError, ValueError):
                pass
        if u.get("line_mode") in ("exact", "multiple"):
            item["line_mode"] = u["line_mode"]
        if u.get("line_value") is not None:
            try:
                item["line_value"] = float(u["line_value"])
            except (TypeError, ValueError):
                pass
        out[k] = item
    return out


def assemble(materials, output_path, cfg=None, options=None, fmt=None):
    """
    materials: list of dict（由 analyze 产出）：
        {title, unit, date, dx, doc, blocks, paras}
        blocks 为有序的源块元素（段落/表格），已排除标题/单位/日期等。
    cfg: {cover_title, date_text, cover_title_size}
    options: {remove_signoff, each_material_new_page}
    fmt: 界面可调格式 {toc|body|h1|h2: {size, line_mode, line_value}}，
        仅作用于 目录内容/正文/一级标题/二级标题；表格/图片不受影响。
    返回输出路径。
    """
    cfg = cfg or {}
    options = options or {}
    fmt = _merge_fmt(fmt)

    ctx = AssemblyCtx()
    document = Document()
    _enable_even_odd(document)
    _add_toc_styles(document, toc_size=fmt["toc"]["size"],
                    line_mode=fmt["toc"]["line_mode"],
                    line_value=fmt["toc"]["line_value"])  # 目录条目样式（界面可调）
    _add_title_char_style(document)  # 材料标题字符样式（防止直接字体被复制进目录条目）

    # 第一节：封面（无页码）
    sec_cover = document.sections[0]
    _setup_section(sec_cover)
    build_cover(document, cfg)
    # 封面与目录分页（需求：目录与封面不是同一页）
    _add_page_break(document)
    # 目录（仍属无页码的封面节，但位于新页）
    build_toc(document, materials)

    # 第二节：正文（新页起始，页码从第1页）
    document.add_section(WD_SECTION.NEW_PAGE)
    sec_body = document.sections[1]
    _setup_section(sec_body)
    _set_pg_num_start(sec_body, 1)
    _add_number_footer(sec_body, "footer")        # 奇数页
    _add_number_footer(sec_body, "even_page_footer")  # 偶数页（双页翻面）
    build_body(document, materials, options, ctx, fmt)

    document.save(output_path)
    # 保存后补丁：并入图片媒体与自动编号定义（使序号/图片正常显示）
    _patch_package(output_path, ctx)
    return output_path


if __name__ == "__main__":
    import sys
    # 简单自测：汇编 tests/ 下的样本
    out = sys.argv[1] if len(sys.argv) > 1 else "tests/out_sample.docx"
    mats = []
    d = "tests/samples"
    if os.path.isdir(d):
        for fn in sorted(os.listdir(d)):
            if fn.lower().endswith(".docx") and not fn.startswith("out_"):
                dx = os.path.join(d, fn)
                info = analyze(dx)
                mats.append({"title": info["title"], "unit": info["unit"],
                             "dx": dx, "blocks": info["blocks"]})
    cfg = {"cover_title": "全市服务业大会",
           "date_text": "2026年8月"}
    assemble(mats, out, cfg=cfg, options={"remove_signoff": True})
    print("assembled ->", out, "materials:", len(mats))
