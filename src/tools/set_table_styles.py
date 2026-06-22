from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

def set_table_borders_ultimate(table, outer_size=1.0, inner_size=0.5, color="000000"):
    """
    终极表格边框设置：解决内外边框不一致、不显示的问题
    :param table: docx 表格对象
    :param outer_size: 外边框宽度 (磅)，默认 1.0pt
    :param inner_size: 内边框宽度 (磅)，默认 0.5pt
    :param color: 16进制颜色字符串 (如 '000000' 为黑色)
    """
    # --- 辅助函数：创建边框 XML 节点 ---
    def make_border_elem(tag, size_pt, color_hex):
        """
        创建一个 w:top/bottom/left/right 元素
        size_pt: 磅值
        """
        elem = OxmlElement(tag)
        # val="single" 表示实线，也可以设为 "double", "dashed" 等
        elem.set(qn('w:val'), 'single')
        # sz 单位是 1/8 pt。例如 0.5pt -> 4
        elem.set(qn('w:sz'), str(int(size_pt * 8)))
        elem.set(qn('w:color'), color_hex)
        elem.set(qn('w:space'), '0')
        return elem

    tbl = table._tbl

    # 1. 获取或创建 tblPr
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    # 2. 修复 tblLook：防止 Word 默认样式隐藏内部边框
    # 这一步非常关键，很多时候边框不显示就是因为这个属性
    tblLook = tblPr.find(qn('w:tblLook'))
    if tblLook is not None:
        tblPr.remove(tblLook)
    new_tblLook = OxmlElement('w:tblLook')
    # 04A0 表示应用所有默认样式，确保我们的设置能生效
    new_tblLook.set(qn('w:val'), '04A0')
    tblPr.append(new_tblLook)

    # 3. 设置表格级边框 (主要控制外框)
    tblBorders = OxmlElement('w:tblBorders')
    for tag in ['w:top', 'w:left', 'w:bottom', 'w:right']:
        tblBorders.append(make_border_elem(tag, outer_size, color))
    # 内部横线和竖线也在这里定义一次作为兜底
    tblBorders.append(make_border_elem('w:insideH', inner_size, color))
    tblBorders.append(make_border_elem('w:insideV', inner_size, color))

    # 移除旧的并添加新的
    old_tblBorders = tblPr.find(qn('w:tblBorders'))
    if old_tblBorders is not None:
        tblPr.remove(old_tblBorders)
    tblPr.append(tblBorders)

    # 4. 【关键】遍历每个单元格，强制设置单元格级边框
    # 很多情况下，单元格级别的属性会覆盖表格级别的属性，导致内部没线
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = OxmlElement('w:tcPr')
                tc.insert(0, tcPr)

            tcBorders = OxmlElement('w:tcBorders')
            # 单元格的四个边都设为内部线条样式
            for tag in ['w:top', 'w:left', 'w:bottom', 'w:right']:
                tcBorders.append(make_border_elem(tag, inner_size, color))

            # 移除旧的并添加新的
            old_tcBorders = tcPr.find(qn('w:tcBorders'))
            if old_tcBorders is not None:
                tcPr.remove(old_tcBorders)
            tcPr.append(tcBorders)


def set_cell_margin_font_line_height(table):
    for i, row in enumerate(table.rows):
        tr = row._tr
        trPr = tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            tr.insert(0, trPr)

        # 设置行高自适应 (atLeast)
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), '400')  # 最小高度约 0.3英寸
        trHeight.set(qn('w:hRule'), 'atLeast') # 关键：允许自动撑开

        old_trHeight = trPr.find(qn('w:trHeight'))
        if old_trHeight is not None:
            trPr.remove(old_trHeight)
        trPr.append(trHeight)

        # 处理第一行背景色
        # if i == 0:
        #     for cell in row.cells:
        #         tc = cell._tc
        #         tcPr = tc.find(qn('w:tcPr'))
        #         if tcPr is None:
        #             tcPr = OxmlElement('w:tcPr')
        #             tc.insert(0, tcPr)

        #         shd = OxmlElement('w:shd')
        #         shd.set(qn('w:val'), 'clear')
        #         shd.set(qn('w:color'), 'auto')
        #         shd.set(qn('w:fill'), '4F81BD') # 经典的 Word 蓝表头颜色

        #         old_shd = tcPr.find(qn('w:shd'))
        #         if old_shd is not None:
        #             tcPr.remove(old_shd)
        #         tcPr.append(shd)

        # 处理单元格文字对齐
        for cell in row.cells:
            # 垂直居中
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            for paragraph in cell.paragraphs:
                # 两端对齐
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                # 设置1.5倍行距
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                
                # 段前0，段后0
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                
                # 禁用对齐到网格
                paragraph.paragraph_format.snap_to_grid = False

                # 统一字号
                for run in paragraph.runs:
                    run.font.size = Pt(10.5) # 五号字

def format_table_complete(table):
    """
    将表格设置为嵌入型（即取消文字环绕/浮动）
    """
    # 1. 获取表格的 XML 根节点 (tbl)
    tbl = table._tbl

    # 2. 查找或创建 tblPr (表格属性节点)
    # python-docx 没有 get_or_add_tblPr，我们需要手动处理
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = etree.SubElement(tbl, qn('w:tblPr'))

    # 3. 查找 tblpPr (表格位置属性，这是导致表格变成"浮动型"的元凶)
    tblpPr = tblPr.find(qn('w:tblpPr'))

    # 4. 如果找到了 tblpPr，说明它是浮动的，必须删除它才能变回嵌入型
    if tblpPr is not None:
        tblPr.remove(tblpPr)

    # --- A. 布局与尺寸 ---
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    # 设置宽度 100%
    old_tblW = tblPr.find(qn('w:tblW'))
    if old_tblW is not None:
        tblPr.remove(old_tblW)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)

    # 设置单元格边距 (Padding) - 避免过大导致看起来像合并
    tblCellMar = OxmlElement('w:tblCellMar')
    for tag in ['w:top', 'w:left', 'w:bottom', 'w:right']:
        margin = OxmlElement(tag)
        margin.set(qn('w:w'), '54')  # 约 0.04英寸，适中
        margin.set(qn('w:type'), 'dxa')
        tblCellMar.append(margin)

    old_tblCellMar = tblPr.find(qn('w:tblCellMar'))
    if old_tblCellMar is not None:
        tblPr.remove(old_tblCellMar)
    tblPr.append(tblCellMar)

    # --- B. 调用上面的终极边框函数 ---
    # 这里设置：外框 1pt，内框 0.5pt，黑色
    set_table_borders_ultimate(table, outer_size=1.0, inner_size=0.5, color="000000")

    # --- C. 遍历行和单元格处理内容与行高 ---
    set_cell_margin_font_line_height(table)

def set_table_styles(doc):
    for table in doc.tables:
        format_table_complete(table)