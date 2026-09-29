# -*- coding: utf-8 -*-
"""
pipeline.py —— 组装流程编排
串联：源文件(.doc/.docx/.wps) -> 转换docx -> 汇编 -> (预览PDF页图)
"""
import os
import re
from . import converter, assembler, preview


def parse_filename(filename):
    """从文件名解析 (单位, 标题)。约定格式：某某单位-标题。"""
    name = os.path.splitext(os.path.basename(filename))[0]
    if "-" in name:
        unit, title = name.split("-", 1)
        return unit.strip(), title.strip()
    return "", name.strip()


def _unit_from_filename(path):
    """单位缺失时的文件名回退：去扩展名、去前导日期。"""
    bn = os.path.splitext(os.path.basename(path))[0]
    bn = re.sub(r"^(\d{8}|20\d{2}年\d{1,2}月\d{1,2}日?|20\d{2}年\d{1,2}月)", "", bn).strip(" -")
    return bn


def analyze_source(path, work_dir):
    """解析单份源文档，返回 {title, unit, date, paras}（标题/单位/日期从内容识别）。"""
    dx = converter.convert_to_docx(path, work_dir)
    info = assembler.analyze(dx)
    if not info["unit"]:
        fb = _unit_from_filename(path)
        if fb:
            info["unit"] = fb
    return info


def run_assemble(sources, output_path, cfg, options, work_dir, fmt=None):
    """
    sources: list of dict {path, unit, title}（unit/title 为空时从内容自动识别）
    cfg: {cover_title, date_text, cover_title_size}
    options: {remove_signoff, each_material_new_page}
    fmt: 界面可调格式 {toc|body|h1|h2: {size, line_mode, line_value}}（可选）
    返回最终 docx 路径。.doc/.wps 由 WPS 转 docx，.docx 直接（或规范化后）汇编。
    """
    os.makedirs(work_dir, exist_ok=True)
    mats = []
    for s in sources:
        dx = converter.convert_to_docx(s["path"], work_dir)
        info = assembler.analyze(dx)
        title = s.get("title") or info["title"]
        unit = s.get("unit") or info["unit"]
        if not unit:
            fb = _unit_from_filename(s["path"])
            if fb:
                unit = fb
        mats.append({"title": title, "unit": unit,
                     "dx": dx, "blocks": info["blocks"]})
    assembler.assemble(mats, output_path, cfg=cfg, options=options, fmt=fmt)
    return output_path


def run_preview(docx_path, work_dir):
    """返回 PNG 页图路径列表（用于预览）。依赖 WPS 转 PDF。"""
    return preview.build_preview(docx_path, work_dir)
