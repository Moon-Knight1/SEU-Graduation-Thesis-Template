# -*- coding: utf-8 -*-
"""
答辩PPT生成脚本
基于东南大学模板创建答辩PPT
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from copy import deepcopy
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

def create_cover_slide(prs):
    """创建封面页

    模板的布局0不含形状，因此通过复制模板第0页（已有全部封面形状）来创建封面。
    形状索引基于模板分析结果：
      1  - 中文标题
      8  - 英文标题（AUTO_SHAPE）
      11 - 答辩学生
      13 - 专业班级
      15 - 指导老师
    """
    # 保存模板封面页引用及其形状副本
    src_slide = prs.slides[0]
    src_shapes_xml = [deepcopy(shape._element) for shape in src_slide.shapes]

    # 清除所有模板幻灯片
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].get(
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        )
        prs.part.drop_rel(rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

    # 使用空白布局创建新幻灯片，再将封面形状复制进来
    slide_layout = prs.slide_layouts[7]  # 空白布局
    slide = prs.slides.add_slide(slide_layout)
    for shape_xml in src_shapes_xml:
        slide.shapes._spTree.append(shape_xml)

    # 设置论文标题
    title = slide.shapes[1]  # 标题文本框
    title.text_frame.text = "基于B/S架构的企业数字化运营平台设计与实现"

    # 设置英文标题
    subtitle = slide.shapes[8]  # 英文标题文本框
    subtitle.text_frame.text = "Design and Implementation of Enterprise Digital Operation Platform Based on B/S Architecture"

    # 设置答辩学生
    student = slide.shapes[11]
    student.text_frame.text = "答辩学生：XXX"

    # 设置专业班级
    major = slide.shapes[13]
    major.text_frame.text = "专业班级：软件工程XX班"

    # 设置指导老师
    teacher = slide.shapes[15]
    teacher.text_frame.text = "指导老师：XXX"

    return slide

if __name__ == "__main__":
    prs = load_template()

    # 创建封面页（会清除模板中的示例幻灯片）
    create_cover_slide(prs)

    save_ppt(prs, OUTPUT_PATH)
