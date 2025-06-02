import os
from time import sleep
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_LINE_SPACING
from docx2pdf import convert
from pypdf import PdfReader

def contains_picture(para):
    # 判断段落是否包含图片（w:drawing）
    for run in para.runs:
        if run.element.xpath('.//w:drawing'):
            return True
    return False

def set_landscape_a4(section):
    section.orientation = WD_ORIENT.LANDSCAPE
    # A4尺寸 21cm x 29.7cm 转换英寸
    section.page_width = Inches(11.69)  # 29.7cm
    section.page_height = Inches(8.27)  # 21cm
    # 边距设置0.5英寸
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

def set_column_count(section, num_columns):
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        cols[0].set('num', str(num_columns))

def resize_images(doc, max_width_inches):
    for shape in doc.inline_shapes:
        width_inches = shape.width.inches
        if width_inches > max_width_inches:
            ratio = max_width_inches / width_inches
            shape.width = int(shape.width * ratio)
            shape.height = int(shape.height * ratio)

def compress_layout(doc_path, output_path, strategy_level, column_count):
    doc = Document(doc_path)

    # 页面设置（所有节都设置）
    for section in doc.sections:
        if strategy_level >= 3:
            set_landscape_a4(section)
        if strategy_level >= 4:
            set_column_count(section, column_count)
    # 计算每栏宽度 = (页宽 - 左右边距) / 栏数，单位英寸
    # 只在分栏≥1时计算
    if column_count >= 1:
        sec = doc.sections[0]
        page_width = sec.page_width.inches
        left_margin = sec.left_margin.inches
        right_margin = sec.right_margin.inches
        available_width = page_width - left_margin - right_margin
        per_column_width = available_width / column_count
    else:
        # 默认为页面内容宽度
        sec = doc.sections[0]
        page_width = sec.page_width.inches
        left_margin = sec.left_margin.inches
        right_margin = sec.right_margin.inches
        per_column_width = page_width - left_margin - right_margin

    # 缩放图片宽度不超过单栏宽度
    resize_images(doc, max_width_inches=per_column_width)

    # 处理段落：删除空白行，字体大小限制，行距设置
    paras = list(doc.paragraphs)
    for para in paras:
        # 删除空白行（无文字，无图片）
        if (not para.text.strip()) and (not contains_picture(para)):
            p_element = para._element
            p_element.getparent().remove(p_element)
            continue

        pf = para.paragraph_format
        if strategy_level >= 1:
            pf.space_before = 0
            pf.space_after = 0
            if contains_picture(para):
                # 含图片段落用至少行距，避免图片被挤压
                pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
                pf.line_spacing = Pt(7)  # 你可以调整小一点，7磅试试
            else:
                # 文字段落用固定行距更密集
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

    for _ in range(10):  # 最多等2秒
        if os.path.exists(temp_pdf_path):
            break
        sleep(0.2)

    if not os.path.exists(temp_pdf_path):
        raise FileNotFoundError("PDF 未生成，转换失败")

    reader = PdfReader(temp_pdf_path)
    return len(reader.pages)

def shrink_to_target_pages(input_path, target_pages, output_path="compressed_output.docx"):
    temp_path = input_path
    for level in range(1, 5):           # 1 到 4 的策略等级
        for cols in range(1, 5):        # 1 到 4 栏
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
