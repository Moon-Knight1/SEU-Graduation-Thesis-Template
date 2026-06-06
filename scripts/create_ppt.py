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

def clear_all_slides(prs):
    """清除模板中的所有幻灯片"""
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].get(
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        )
        prs.part.drop_rel(rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

def create_cover_slide(prs, src_shapes_xml=None):
    """创建封面页

    模板的布局0不含形状，因此通过复制模板第0页（已有全部封面形状）来创建封面。
    形状索引基于模板分析结果：
      1  - 中文标题
      8  - 英文标题（AUTO_SHAPE）
      11 - 答辩学生
      13 - 专业班级
      15 - 指导老师

    Args:
        prs: Presentation对象
        src_shapes_xml: 预先复制的形状XML列表，如果为None则从模板第0页获取
    """
    if src_shapes_xml is None:
        src_shapes_xml = [deepcopy(shape._element) for shape in prs.slides[0].shapes]

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

def create_toc_slide(prs, src_shapes_xml=None):
    """创建目录页

    通过复制模板第1页（目录页）来创建目录。
    模板目录页共28个形状，5个章节按垂直位置从上到下排列：
      18,19,20 - 第1章：编号、中文标题、英文标题
       5, 6, 7 - 第2章
       8, 9,10 - 第3章
      11,12,13 - 第4章
      14,15,16 - 第5章
    其他形状：0=背景组, 1=装饰, 2="目录", 3="CONTENTS",
              4=装饰, 17/21-25=分隔线, 26=页脚, 27=图片

    Args:
        prs: Presentation对象
        src_shapes_xml: 预先复制的形状XML列表，如果为None则从模板第1页获取
    """
    # 目录章节数据：(中文标题, 英文标题)
    toc_sections = [
        ("项目概述与系统架构", "Project Overview and Architecture"),
        ("核心功能模块设计", "Core Function Modules"),
        ("技术亮点与创新", "Technical Highlights"),
        ("系统测试与性能优化", "Testing and Optimization"),
        ("总结与展望", "Conclusion and Prospects"),
    ]

    # 各章节在模板形状中的索引：(编号shape, 中文标题shape, 英文标题shape)
    section_indices = [
        (18, 19, 20),  # 第1章 - 最上方
        (5,  6,  7),   # 第2章
        (8,  9,  10),  # 第3章
        (11, 12, 13),  # 第4章
        (14, 15, 16),  # 第5章 - 最下方
    ]

    if src_shapes_xml is None:
        src_shapes_xml = [deepcopy(shape._element) for shape in prs.slides[1].shapes]

    # 使用空白布局创建目录页，再将形状复制进来
    slide_layout = prs.slide_layouts[7]
    slide = prs.slides.add_slide(slide_layout)
    for shape_xml in src_shapes_xml:
        slide.shapes._spTree.append(shape_xml)

    # 更新各章节文本
    for i, (cn_title, en_title) in enumerate(toc_sections):
        num_idx, cn_idx, en_idx = section_indices[i]

        # 设置编号
        num_shape = slide.shapes[num_idx]
        num_shape.text_frame.text = f"0{i + 1}"

        # 设置中文标题
        cn_shape = slide.shapes[cn_idx]
        cn_shape.text_frame.text = cn_title

        # 设置英文标题
        en_shape = slide.shapes[en_idx]
        en_shape.text_frame.text = en_title

    return slide

def create_content_slide(prs, title_cn, title_en, section_num,
                         content_items=None, src_shapes_xml=None):
    """创建内容页（章节标题页）

    通过复制模板第2页（第1章章节标题页）来创建内容页。
    模板章节标题页共11个形状，结构一致：
      0  - 中文标题文本框
      1  - 英文标题文本框
      2  - 章节编号文本框（"01"、"02"等大字）
      3  - 底部装饰线
      4  - 底部装饰矩形
      5  - 顶部装饰线
      6  - 顶部装饰矩形
      7  - 页脚组合
      8  - 页眉文字（"止于至善"）
      9  - 校徽图片
      10 - 右上角装饰组合

    Args:
        prs: Presentation对象
        title_cn: 中文标题字符串
        title_en: 英文标题字符串
        section_num: 章节编号字符串（如 "01"、"02"）
        content_items: 内容条目列表（可选），若提供则创建包含内容文本的文本框
        src_shapes_xml: 预先复制的形状XML列表，如果为None则从模板第2页获取

    Returns:
        创建的幻灯片对象
    """
    if src_shapes_xml is None:
        src_shapes_xml = [deepcopy(shape._element) for shape in prs.slides[2].shapes]

    # 使用空白布局创建新幻灯片，再将形状复制进来
    slide_layout = prs.slide_layouts[7]  # 空白布局
    slide = prs.slides.add_slide(slide_layout)
    for shape_xml in src_shapes_xml:
        slide.shapes._spTree.append(shape_xml)

    # 设置中文标题
    slide.shapes[0].text_frame.text = title_cn

    # 设置英文标题
    slide.shapes[1].text_frame.text = title_en

    # 设置章节编号
    slide.shapes[2].text_frame.text = section_num

    # 如果提供了内容条目，在页面下方创建内容文本框
    if content_items:
        from pptx.util import Inches, Pt
        from pptx.enum.text import PP_ALIGN

        left = Inches(1.0)
        top = Inches(2.0)
        width = Inches(8.0)
        height = Inches(3.5)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True

        for i, item in enumerate(content_items):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = item
            p.font.size = Pt(14)
            p.space_after = Pt(6)

    return slide


if __name__ == "__main__":
    prs = load_template()

    # 在清除前复制模板幻灯片的形状数据
    cover_shapes_xml = [deepcopy(shape._element) for shape in prs.slides[0].shapes]
    toc_shapes_xml = [deepcopy(shape._element) for shape in prs.slides[1].shapes]
    content_shapes_xml = [deepcopy(shape._element) for shape in prs.slides[2].shapes]

    # 清空模板幻灯片
    clear_all_slides(prs)

    # 创建幻灯片
    create_cover_slide(prs, cover_shapes_xml)
    create_toc_slide(prs, toc_shapes_xml)

    # 测试内容页模板函数
    create_content_slide(
        prs,
        title_cn="项目概述与系统架构",
        title_en="Project Overview and Architecture",
        section_num="01",
        src_shapes_xml=content_shapes_xml,
    )

    save_ppt(prs, OUTPUT_PATH)
