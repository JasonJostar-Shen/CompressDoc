import os
import re
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
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

def set_column_count(section, num_columns):
    sectPr = section._sectPr
    # Remove old <w:cols> if it exists
    cols = sectPr.find(qn('w:cols'))
    if cols is not None:
        sectPr.remove(cols)

    # Create new <w:cols>
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), str(num_columns))     # 栏数
    cols.set(qn('w:space'), '0')                # 栏间距为0
    cols.set(qn('w:equalWidth'), '1')           # 强制等宽
    sectPr.append(cols)

def resize_images(doc, max_width_inches):
    for shape in doc.inline_shapes:
        width_inches = shape.width.inches
        if width_inches > max_width_inches:
            ratio = max_width_inches / width_inches
            shape.width = int(shape.width * ratio)
            shape.height = int(shape.height * ratio)

def compress_layout(doc_path, output_path, strategy_level, column_count):
    doc = Document(doc_path)

    for section in doc.sections:
        if strategy_level >= 3:
            set_landscape_a4(section)
        if strategy_level >= 4:
            set_column_count(section, column_count)

    sec = doc.sections[0]
    page_width = sec.page_width.inches
    left_margin = sec.left_margin.inches
    right_margin = sec.right_margin.inches
    available_width = page_width - left_margin - right_margin
    per_column_width = available_width / max(1, column_count)

    resize_images(doc, max_width_inches=per_column_width)

    for para in list(doc.paragraphs):
        if (not para.text.strip()) and (not contains_picture(para)):
            p_element = para._element
            p_element.getparent().remove(p_element)
            continue

        pf = para.paragraph_format
        if strategy_level >= 1:
            pf.space_before = 0
            pf.space_after = 0
            pf.left_indent = Inches(0.01)  # 0.5 个字符宽
            pf.right_indent = Inches(0)
            if contains_picture(para):
                pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
                pf.line_spacing = Pt(7)
            else:
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                pf.line_spacing = Pt(7)

        if strategy_level >= 2:
            for run in para.runs:
                run.style = None  # 脱离样式控制
                text = run.text
                if not text:
                    continue
                if re.search(r'[\u4e00-\u9fff]', text):
                    run.font.size = Pt(6)
                    run.font.name = 'SimSun'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
                else:
                    if not run.font.size or run.font.size.pt > 7:
                        run.font.size = Pt(7)

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
    for level in range(1, 5):
        for cols in range(1, 7):  # 支持1到6栏
            compress_layout(temp_path, output_path, level, cols)
            try:
                pages = get_pdf_page_count(output_path)
            except Exception as e:
                print("❌ PDF 转换失败，请检查 WPS/Word 安装：", e)
                return
            print(f"尝试 Level {level} + {cols} 栏：页数 = {pages}")
            if pages <= target_pages:
                print(f"✅ 成功压缩到 {pages} 页（Level {level} + {cols} 栏）")
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
