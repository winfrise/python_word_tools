import os
from tools.set_page_margin import set_page_margin
from tools.set_heading_level_styles import set_heading_level_style
from tools.convert_font_to_heading import convert_font_to_heading
from tools.set_table_styles import set_table_styles
from tools.set_text_styles import set_text_style
from tools.set_global_font import set_global_font
from tools.set_all_font_styles import set_all_font_styles

def format_document(input_path, output_path):
    from docx import Document
    from docx.shared import Cm, Inches

    # 1. 创建或打开文档
    doc = Document(input_path)

    convert_font_to_heading(doc)

    set_global_font(doc)

    set_all_font_styles(doc)

    set_page_margin(doc)

    set_heading_level_style(doc)

    set_text_style(doc)

    set_table_styles(doc)




    # 4. 保存文档
    doc.save(output_path)




if __name__ == "__main__":
    input_path = "/Users/teacher/Desktop/word——test/家居产品销售手册(1).docx"

    output_path = input_path.replace('.docx', '-formated.docx')

    format_document(input_path, output_path)
