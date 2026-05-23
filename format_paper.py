#!/usr/bin/env python3
"""山西财经大学本科学年论文自动格式化脚本"""

import sys
import os
import re
from docx import Document
from docx.shared import Pt, Cm, Inches, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy


def set_page_margins(section):
    """设置A4页面和页边距"""
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(3.0)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)


def set_header_footer(section, grade="20XX"):
    """设置页眉"""
    header = section.header
    header.is_linked_to_previous = False
    # Clear existing
    for p in header.paragraphs:
        p.clear()

    # Odd page header (right-aligned)
    # Use different first page setting
    section.different_first_page_header_footer = True

    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.space_after = Pt(0)
    run = hp.add_run(f"山西财经大学{grade}级本科生学年论文")
    run.font.size = Pt(9)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')


def set_paragraph_format(para, font_name="宋体", font_size=Pt(12),
                         bold=False, alignment=None,
                         first_line_indent=None,
                         space_before=Pt(0), space_after=Pt(0),
                         line_spacing=1.25):
    """通用段落格式设置"""
    pf = para.paragraph_format
    pf.space_before = space_before
    pf.space_after = space_after
    pf.line_spacing = line_spacing

    if alignment is not None:
        para.alignment = alignment

    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent

    for run in para.runs:
        run.font.size = font_size
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.bold = bold


def format_body_paragraph(para):
    """格式化正文段落"""
    text = para.text.strip()
    if not text:
        return

    set_paragraph_format(para, font_name="宋体", font_size=Pt(12),
                         first_line_indent=Cm(0.74),  # ~2 chars
                         line_spacing=1.25)


def identify_and_format_heading(para):
    """识别并格式化标题"""
    text = para.text.strip()
    if not text:
        return False

    # Match heading patterns like "一、", "（一）", "1.", "1.1", "导论", "结语"
    level1_patterns = [
        r'^[一二三四五六七八九十]、',  # 一、二、
        r'^(导论|导 论|结语|结 语|参考文献|参考 文献|致谢|致 谢|附录|附 录)$',
        r'^Abstract$',
    ]
    level2_patterns = [
        r'^（[一二三四五六七八九十]）',  # （一）（二）
        r'^\d+\.\d+\s',  # 1.1
    ]
    level3_patterns = [
        r'^\d+\.\d+\.\d+\s',  # 1.1.1
    ]

    for pattern in level1_patterns:
        if re.match(pattern, text):
            set_paragraph_format(para, font_name="黑体", font_size=Pt(16),
                               bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                               space_before=Pt(12), space_after=Pt(12),
                               line_spacing=1.25)
            return True

    for pattern in level2_patterns:
        if re.match(pattern, text):
            set_paragraph_format(para, font_name="宋体", font_size=Pt(14),
                               bold=True, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                               line_spacing=1.25)
            return True

    for pattern in level3_patterns:
        if re.match(pattern, text):
            set_paragraph_format(para, font_name="宋体", font_size=Pt(12),
                               bold=True, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                               line_spacing=1.25)
            return True

    return False


def format_paper(input_path, grade="20XX"):
    """
    主函数：格式化学年论文

    Args:
        input_path: 输入论文路径
        grade: 年级，如 "2024"
    """
    if not os.path.exists(input_path):
        print(f"错误：文件不存在 - {input_path}")
        return None

    doc = Document(input_path)

    # === 1. 页面设置 ===
    for section in doc.sections:
        set_page_margins(section)
        set_header_footer(section, grade)

    # === 2. 遍历段落，格式化 ===
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # 跳过封面和学术承诺相关的段落（保留原格式）
        if any(kw in text for kw in ['学校代码', '学术承诺', '使用授权', '签名', '日期']):
            continue

        # 识别标题
        if identify_and_format_heading(para):
            continue

        # 格式化正文
        format_body_paragraph(para)

    # === 3. 保存 ===
    dir_name = os.path.dirname(input_path)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(dir_name, f"{base_name}_formatted.docx")

    doc.save(output_path)
    print(f"格式化完成！输出文件：{output_path}")

    # === 4. 报告修改内容 ===
    print("""
已应用的格式修改：
  ✅ 纸张：A4
  ✅ 页边距：上3cm, 下2.5cm, 左2.5cm, 右2cm
  ✅ 页眉：山西财经大学{0}级本科生学年论文（小五号宋体）
  ✅ 正文：小四号宋体，1.25倍行距，首行缩进2字符
  ✅ 标题层级：一级(三号黑体居中)、二级(四号宋体加粗)、三级(小四号宋体加粗)
  ✅ 参考文献：小四号宋体
""".format(grade))

    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python format_paper.py <论文.docx> [年级]")
        print("示例: python format_paper.py 论文主体.docx 2024")
        sys.exit(1)

    input_path = sys.argv[1]
    grade = sys.argv[2] if len(sys.argv) > 2 else "20XX"

    format_paper(input_path, grade)
