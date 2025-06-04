import os
from time import sleep
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_LINE_SPACING
from docx2pdf import convert
from pypdf import PdfReader
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def contains_picture(para):
    for run in para.runs:
        if run.element.xpath('.//w:drawing'):
            return True
    return False


def set_landscape_a4(section):
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)  # 29.7cm
    section.page_height = Inches(8.27)  # 21cm
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)


def set_column_count(section, num_columns, space_twips=720):
    sectPr = section._sectPr
    cols = sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols')
        sectPr.append(cols)
    cols.set(qn('w:num'), str(num_columns))
    cols.set(qn('w:space'), str(space_twips))  # 栏间距，默认720 Twips = 0.5英寸


def resize_images(doc, max_width_inches):
    for shape in doc.inline_shapes:
        width_inches = shape.width.inches
        if width_inches > max_width_inches:
            ratio = max_width_inches / width_inches
            shape.width = int(shape.width * ratio)
            shape.height = int(shape.height * ratio)


def compress_layout(doc_path, output_path, strategy_level, column_count):
    doc = Document(doc_path)

    # 设置分栏（>=1时设置，不然移除分栏设置）
    for section in doc.sections:
        if column_count >= 1:
            set_column_count(section, column_count)
        else:
            sectPr = section._sectPr
            cols = sectPr.find(qn('w:cols'))
            if cols is not None:
                sectPr.remove(cols)

    # 高级策略调整页面方向和边距
    if strategy_level >= 3:
        for section in doc.sections:
            set_landscape_a4(section)

    # 计算单栏宽度
    sec = doc.sections[0]
    page_width = sec.page_width.inches
    left_margin = sec.left_margin.inches
    right_margin = sec.right_margin.inches
    available_width = page_width - left_margin - right_margin
    per_column_width = available_width / max(column_count, 1)

    resize_images(doc, max_width_inches=per_column_width)

    paras = list(doc.paragraphs)
    for para in paras:
        if (not para.text.strip()) and (not contains_picture(para)):
            p_element = para._element
            p_element.getparent().remove(p_element)
            continue

        pf = para.paragraph_format
        if strategy_level >= 1:
            pf.space_before = 0
            pf.space_after = 0
            if contains_picture(para):
                pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
                pf.line_spacing = Pt(7)
            else:
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                pf.line_spacing = Pt(7)

        if strategy_level >= 2:
            for run in para.runs:
                if run.font.size and run.font.size.pt > 8:
                    run.font.size = Pt(8)

    doc.save(output_path)


def get_pdf_page_count(docx_path):
    temp_pdf_path = "temp_output.pdf"
    try:
        convert(docx_path, temp_pdf_path)
    except Exception as e:
        print("⚠️ 转换时报错（可能是假报错）：", e)

    for _ in range(10):
        if os.path.exists(temp_pdf_path):
            break
        sleep(0.2)

    if not os.path.exists(temp_pdf_path):
        raise FileNotFoundError("PDF 未生成，转换失败")

    reader = PdfReader(temp_pdf_path)
    return len(reader.pages)


def shrink_to_target_pages(input_path, target_pages, output_path="compressed_output.docx"):
    temp_path = input_path
    for cols in range(1, 5):          # 先尝试不同分栏数 1-4栏
        for level in range(1, 5):     # 再尝试不同策略等级 1-4级
            compress_layout(temp_path, output_path, level, cols)
            try:
                pages = get_pdf_page_count(output_path)
            except Exception as e:
                print("❌ PDF 转换失败，请检查 WPS/Word 安装：", e)
                return
            print(f"尝试 分栏 {cols} 栏 + 策略等级 {level} ：页数 = {pages}")
            if pages <= target_pages:
                print(f"✅ 成功压缩到 {pages} 页（分栏 {cols} 栏 + 策略等级 {level}）")
                return
            temp_path = output_path
    print("❌ 所有压缩策略尝试后仍无法达到目标页数。")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("用法: python word_shrinker_wps.py 输入文件.docx 目标页数")
    else:
        input_doc = sys.argv[1]
        target = int(sys.argv[2])
        shrink_to_target_pages(input_doc, target, "compressed_output.docx")
