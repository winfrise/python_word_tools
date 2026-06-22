from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

def set_global_font(doc):
    # 修改全局默认正文样式
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'  # 英文字体
    style.font.size = Pt(12)             # 小四号字
    # 必须同时设置东亚字体
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # --- 新增：颜色设置 ---
    # 使用 RGBColor(r, g, b) 定义颜色
    # 示例：黑色 (0,0,0)，深灰色 (51,51,51)，红色 (255,0,0)
    target_color = RGBColor(0, 0, 0)

    # 将颜色应用到样式的字体对象上
    style.font.color.rgb = target_color