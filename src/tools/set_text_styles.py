from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_text_style(doc):
    for para in doc.paragraphs:
        # 跳过标题样式的段落
        if para.style.name.startswith('Heading') or para.style.name.startswith('标题'):
            continue  

        # 设置1.5倍行距
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE 
        
        # 关闭“与网格对齐”
        para.paragraph_format.snap_to_grid = False

        # 对剩余的纯正文段落进行字体修改
        for run in para.runs:
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(0, 0, 0) # 设置为黑色