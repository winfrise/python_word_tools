from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn


def set_heading_level_style(doc):

    # 1. 定义不同级别标题的样式配置字典
    # 您可以根据实际需求随时修改这里的参数
    heading_styles_config = {
        1: {"cn_font": "微软雅黑", "en_font": "Arial", "size": Pt(22), "color": RGBColor(0, 51, 102), "bold": True, "align": WD_PARAGRAPH_ALIGNMENT.CENTER},
        2: {"cn_font": "微软雅黑", "en_font": "Arial", "size": Pt(18), "color": RGBColor(0, 102, 153), "bold": True, "align": WD_PARAGRAPH_ALIGNMENT.LEFT},
        3: {"cn_font": "黑体", "en_font": "Times New Roman", "size": Pt(16), "color": RGBColor(0, 0, 0), "bold": True, "align": WD_PARAGRAPH_ALIGNMENT.LEFT},
        4: {"cn_font": "宋体", "en_font": "Times New Roman", "size": Pt(14), "color": RGBColor(0, 0, 0), "bold": True, "align": WD_PARAGRAPH_ALIGNMENT.LEFT},
        5: {"cn_font": "宋体", "en_font": "Times New Roman", "size": Pt(13), "color": RGBColor(0, 0, 0), "bold": True, "align": WD_PARAGRAPH_ALIGNMENT.LEFT},
        6: {"cn_font": "宋体", "en_font": "Times New Roman", "size": Pt(12), "color": RGBColor(50, 50, 50), "bold": True, "align": WD_PARAGRAPH_ALIGNMENT.LEFT},
        7: {"cn_font": "宋体", "en_font": "Times New Roman", "size": Pt(12), "color": RGBColor(50, 50, 50), "bold": False, "align": WD_PARAGRAPH_ALIGNMENT.LEFT},
        8: {"cn_font": "宋体", "en_font": "Times New Roman", "size": Pt(12), "color": RGBColor(50, 50, 50), "bold": False, "align": WD_PARAGRAPH_ALIGNMENT.LEFT},
        9: {"cn_font": "宋体", "en_font": "Times New Roman", "size": Pt(12), "color": RGBColor(50, 50, 50), "bold": False, "align": WD_PARAGRAPH_ALIGNMENT.LEFT},
    }
    # 遍历段落，根据级别应用差异化样式
    for para in doc.paragraphs:
        # 跳过不是标题的段落
        if not para.style.name.startswith(('Heading', '标题')):
            continue

        # 提取标题的级别数字（例如 'Heading 1' -> 1）
        try:
            level = int(para.style.name.split()[-1])
        except ValueError:
            continue  # 如果标题没有数字，跳过
            
        # 从配置字典中获取当前级别的样式，如果字典中没有该级别，则跳过
        config = heading_styles_config.get(level)
        if not config:
            continue

        # 设置段落对齐方式
        para.alignment = config["align"]
        
        # 遍历段落内的所有 Run 进行字符级格式化
        for run in para.runs:
            # 英文字体
            run.font.name = config["en_font"]
            # 中文字体（底层 XML 设置）
            run._element.rPr.rFonts.set(qn('w:eastAsia'), config["cn_font"])
            # 字号
            run.font.size = config["size"]
            # 颜色
            run.font.color.rgb = config["color"]
            # 加粗
            run.font.bold = config["bold"]

    return doc
