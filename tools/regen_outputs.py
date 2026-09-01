# -*- coding: utf-8 -*-
"""重新生成 测试/汇编结果.docx 与 测试2/汇编结果（测试2）.docx（保持原材料顺序）。"""
import os, sys, glob
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from app.core import pipeline

ROOT = os.path.join(os.path.dirname(__file__), "..")
os.chdir(ROOT)

CFG = {"cover_title": "全市服务业大会",
       "date_text": "2024年"}
OPTIONS = {"remove_signoff": True, "each_material_new_page": True}

T2_ORDER = [
    "交运-2024年上半年全区交通运输经济运行情况（区交运局）.docx",
    "住建-住建局关于经济分析情况汇报20240724.doc",
    "农业-2024年上半年农林牧渔经济运行情况分析报告(2).docx",
    "工科-2024.7.26上半年工贸经济运行情况汇报（四版）.docx",
    "百千万-0725区“百千万工程”指挥办关于“百千万工程”进展情况、存在问题及下一步工作打算.doc",
    "自然资源-20240722自然资源局2024年财税收入情况汇报.docx",
    "税务-税务局2024年上半年税收情况分析（上半年经济调度会）.doc",
    "统计-2024年上半年主要经济指标完成情况及长短板分析.doc",
    "财政-7.23关于今年以来财政收支及重点收入推进情况的报告.wps",
]

def build(test_dir, order_names, out_name):
    sources = []
    for name in order_names:
        p = os.path.join(test_dir, name)
        assert os.path.isfile(p), p
        sources.append({"path": p, "unit": "", "title": ""})
    out = os.path.join(test_dir, out_name)
    pipeline.run_assemble(sources, out, CFG, OPTIONS, "_regen_work")
    print("built:", out)

if __name__ == "__main__":
    build("测试2", T2_ORDER, "汇编结果（测试2）.docx")
    CFG["date_text"] = "2026年8月"   # 测试 组的封面日期原为 2026年8月
    build("测试", sorted(os.path.basename(f) for f in glob.glob("测试/*.docx")
                         if "汇编结果" not in f), "汇编结果.docx")
