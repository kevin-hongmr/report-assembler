# -*- coding: utf-8 -*-
"""
preview.py —— 预览模块
流程：assembler 产出 docx -> WPS 转 PDF -> PyMuPDF 渲染为 PNG 页图。
与 PyQt 解耦：本模块只产出 PNG 文件列表，由 UI 加载。
"""
import os
from .converter import convert_to_pdf


def _import_fitz():
    """兼容 PyMuPDF 1.24+（推荐 import pymupdf）与旧版（import fitz）。"""
    try:
        import pymupdf as fitz  # PyMuPDF >= 1.24
    except Exception:
        import fitz  # 旧版
    return fitz


def has_pymupdf():
    """预览渲染是否可用（PyMuPDF 是否已安装）。"""
    try:
        _import_fitz()
        return True
    except Exception:
        return False


def render_pdf_to_images(pdf_path, out_dir, scale=1.6):
    """把 PDF 每页渲染为 PNG，返回 PNG 路径列表（按页序）。"""
    try:
        fitz = _import_fitz()
    except Exception:
        raise RuntimeError(
            "预览渲染需要 PyMuPDF 库。请在该环境执行：\n"
            "  app\\.venv\\Scripts\\python.exe -m pip install PyMuPDF\n"
            "然后重新点击「预览」。"
        )
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    paths = []
    mat = fitz.Matrix(scale, scale)
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        p = os.path.join(out_dir, f"page_{i+1}.png")
        pix.save(p)
        paths.append(p)
    doc.close()
    return paths


def build_preview(docx_path, work_dir):
    """
    端到端：docx -> pdf -> 页图。返回 PNG 路径列表。
    work_dir: 临时工作目录（建议每个会话独立）。
    依赖 WPS 将 docx 转为 PDF（需 WPS 专业版/政府版支持命令行转换）。
    """
    pdf = convert_to_pdf(docx_path, work_dir)
    return render_pdf_to_images(pdf, work_dir)
