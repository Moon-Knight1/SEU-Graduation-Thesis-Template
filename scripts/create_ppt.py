# -*- coding: utf-8 -*-
"""
答辩PPT生成脚本
基于东南大学模板创建答辩PPT
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import os

# 模板路径
TEMPLATE_PATH = r"E:\SEU-Graduation-Thesis-Template\doc\SEU01东南大学—东南集市.pptx"
OUTPUT_PATH = r"E:\SEU-Graduation-Thesis-Template\output\答辩PPT.pptx"

def load_template():
    """加载PPT模板"""
    prs = Presentation(TEMPLATE_PATH)
    return prs

def save_ppt(prs, path):
    """保存PPT文件"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    prs.save(path)
    print(f"PPT已保存到: {path}")

if __name__ == "__main__":
    prs = load_template()
    print(f"模板已加载，共 {len(prs.slides)} 页")
    print(f"可用布局: {len(prs.slide_layouts)} 种")
    for i, layout in enumerate(prs.slide_layouts):
        print(f"  布局 {i}: {layout.name}")
