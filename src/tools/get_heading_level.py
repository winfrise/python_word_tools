from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


def get_heading_level(style_name):
    """
    从样式名中提取标题级别数字。
    例如：从 "标题 1" 或 "Heading 1" 中提取出 1。
    如果不是标题，返回 None。
    """
    if "标题" in style_name:
        parts = style_name.split("标题")
    elif "Heading" in style_name:
        parts = style_name.split("Heading")
    else:
        return None
    
    # 检查分割后的第二部分是否是数字
    if len(parts) > 1 and parts[1].strip().isdigit():
        return int(parts[1].strip())
    return None

heading_styles = {
    1: { # 标题 1 的样式
        'font_size': Pt(20),
        'font_color': RGBColor(2, 2, 2), # 黑色
        'bg_color': None,                # 无背景色
        'alignment': WD_ALIGN_PARAGRAPH.CENTER, # 居中
        'first_line_indent': None        # 无缩进
    },
    2: { # 标题 2 的样式
        'font_size': Pt(16),
        'font_color': RGBColor(255, 0, 00), # 深灰色
        'bg_color': None,
        'alignment': WD_ALIGN_PARAGRAPH.LEFT, # 左对齐
        'first_line_indent': None
    },
    3: { # 标题 3 的样式
        'font_size': Pt(12),
        'font_color': RGBColor(100, 100, 100), # 灰色
        'bg_color': RGBColor(240, 240, 240), # 浅灰色背景
        'alignment': WD_ALIGN_PARAGRAPH.LEFT,
        'first_line_indent': Pt(24) # 首行缩进2字符
    },
    4: {
        'font_size': Pt(12),
        'font_color': RGBColor(100, 100, 100), # 灰色
        'bg_color': RGBColor(240, 240, 240), # 浅灰色背景
        'alignment': WD_ALIGN_PARAGRAPH.LEFT,
        'first_line_indent': Pt(24) # 首行缩进2字符
    },
    5: {

    },
    6: {

    },
    7: {

    },
    8: {

    },
    9: {

    }
}