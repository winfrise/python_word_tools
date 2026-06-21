from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_table_width_to_100_percent(table):
    """将指定表格的宽度设置为100%"""
    tblPr = table._tbl.tblPr
    # 如果 tblPr 不存在，则创建
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
    
    # 查找或创建 w:tblW 元素
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    
    # 设置宽度类型为百分比，值为5000（即100%）
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')

def set_table_to_embedded(table):
    """将指定表格设置为嵌入型"""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)

    # 查找或创建 w:tblpPr 元素
    tblpPr = tblPr.find(qn('w:tblpPr'))
    if tblpPr is None:
        tblpPr = OxmlElement('w:tblpPr')
        tblPr.append(tblpPr)

    # 设置垂直锚点为文本（嵌入型）
    tblpPr.set(qn('w:vertAnchor'), 'text')
    # 设置水平锚点为页边距（嵌入型）
    tblpPr.set(qn('w:horzAnchor'), 'margin')
    # 可选：设置左缩进为0，确保对齐
    tblpPr.set(qn('w:leftFromText'), '0')
    tblpPr.set(qn('w:rightFromText'), '0')
    tblpPr.set(qn('w:topFromText'), '0')
    tblpPr.set(qn('w:bottomFromText'), '0')

def set_cell_margins(cell, top=50, bottom=50, left=100, right=100):
    """
    设置单元格内边距
    :param cell: 单元格对象
    :param top: 上边距（缇）
    :param bottom: 下边距（缇）
    :param left: 左边距（缇）
    :param right: 右边距（缇）
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # 查找或创建 w:tcMar 元素
    tcMar = tcPr.find(qn('w:tcMar'))
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    
    # 设置各方向边距
    for direction, value in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        margin_elem = tcMar.find(qn(f'w:{direction}'))
        if margin_elem is None:
            margin_elem = OxmlElement(f'w:{direction}')
            tcMar.append(margin_elem)
        margin_elem.set(qn('w:w'), str(value))
        margin_elem.set(qn('w:type'), 'dxa')

def set_cell_format(cell):
    font_name = None
    font_size = None
    line_spacing_pt = None

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # --- 1. 设置垂直居中 (保留你原有的功能) ---
    vAlign = tcPr.find(qn('w:vAlign'))
    if vAlign is None:
        vAlign = OxmlElement('w:vAlign')
        tcPr.append(vAlign)
    vAlign.set(qn('w:val'), 'center')

    # --- 2. 遍历单元格内的所有段落进行格式化 ---
    for paragraph in cell.paragraphs:
        p = paragraph._p
        pPr = p.get_or_add_pPr()

        # A. 设置行距 (Line Spacing)
        if line_spacing_pt is not None:
            spacing = pPr.find(qn('w:spacing'))
            if spacing is None:
                spacing = OxmlElement('w:spacing')
                pPr.append(spacing)

            # Word 内部行距单位是 twip (1/20 point)，所以 pt * 20
            # w:lineRule="exact" 表示固定值行距
            spacing.set(qn('w:line'), str(int(line_spacing_pt * 20)))
            spacing.set(qn('w:lineRule'), 'exact')

        # B. 遍历段落中的 Run (文本块) 设置字体和字号
        for run in paragraph.runs:
            rPr = run._r.get_or_add_rPr()

            # 设置字体 (Font Name)
            if font_name is not None:
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = OxmlElement('w:rFonts')
                    rPr.append(rFonts)

                # 同时设置 ascii (西文) 和 eastAsia (中文) 字体
                rFonts.set(qn('w:ascii'), font_name)
                rFonts.set(qn('w:hAnsi'), font_name)
                rFonts.set(qn('w:eastAsia'), font_name)

            # 设置字号 (Font Size)
            if font_size is not None:
                sz = rPr.find(qn('w:sz'))
                if sz is None:
                    sz = OxmlElement('w:sz')
                    rPr.append(sz)

                # Word 内部字号单位是 half-point (半磅)，所以 pt * 2
                sz.set(qn('w:val'), str(int(font_size * 2)))


def set_table_borders(table, border_color='333333', border_size=1):
    """
    设置表格所有边框的样式（统一风格）
    :param table: 表格对象
    :param border_color: 边框颜色 (十六进制字符串，如 'FF0000' 表示红色)
    :param border_size: 边框粗细 (单位：pt，例如 0.5, 1.0, 1.5)
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    # 查找或创建 w:tblBorders 元素
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)

    # 计算 sz 值 (1pt = 8 sz)
    # Word中常见的边框宽度：0.5pt=4, 0.75pt=6, 1pt=8, 1.5pt=12, 2.25pt=18
    size_val = str(int(border_size * 8))

    # 定义需要设置的边框方向
    borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']

    for border_name in borders:
        # 查找现有的边框元素
        b = tblBorders.find(qn(f'w:{border_name}'))
        if b is None:
            b = OxmlElement(f'w:{border_name}')
            tblBorders.append(b)

        # 设置属性
        b.set(qn('w:val'), 'single')      # 单实线
        b.set(qn('w:sz'), size_val)       # 粗细
        b.set(qn('w:color'), border_color)# 颜色
        b.set(qn('w:space'), '0')         # 边距