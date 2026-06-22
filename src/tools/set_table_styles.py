from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

# 把表格设置为嵌入型
def set_table_inline_center(table):
    """
    将表格设置为嵌入型（即删除所有环绕/定位属性）
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    
    # 查找并删除 w:tblpPr 节点（这是控制环绕的核心节点）
    tblpPr = tblPr.find(qn('w:tblpPr'))
    if tblpPr is not None:
        tblPr.remove(tblpPr)
        
    # 可选：同时删除可能存在的 w:tblInd (缩进)，确保完全重置
    # tblInd = tblPr.find(qn('w:tblInd'))
    # if tblInd is not None:
    #     tblPr.remove(tblInd)

    # 表格整体居中对齐
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

# 设置表格宽度100%
def set_table_full_width(table):
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

# 设置表格内边距
def set_table_padding(table):
    tblCellMar = OxmlElement('w:tblCellMar')
    for tag in ['w:top', 'w:left', 'w:bottom', 'w:right']:
        margin = OxmlElement(tag)
        margin.set(qn('w:w'), '150')  # 边距
        margin.set(qn('w:type'), 'dxa')
        tblCellMar.append(margin)

    tbl = table._tbl
    tblPr = tbl.tblPr
    old_tblCellMar = tblPr.find(qn('w:tblCellMar'))
    if old_tblCellMar is not None:
        tblPr.remove(old_tblCellMar)
    tblPr.append(tblCellMar)

# 设置表格行的高度
def set_table_row_height(table):
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

# 设置表格第一行背景样式
def set_table_first_row_background(table):
    row = table.rows[0]
    if not row:
        return

    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)

        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '4F81BD') # 经典的 Word 蓝表头颜色

        old_shd = tcPr.find(qn('w:shd'))
        if old_shd is not None:
            tcPr.remove(old_shd)
        tcPr.append(shd)

def set_table_borders_ultimate(table):
    outer_size=1.0
    inner_size=0.5
    color="000000"
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


def set_table_cell_align_line_height(table):
    for i, row in enumerate(table.rows):

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

def set_table_styles(doc):
    for table in doc.tables:
        set_table_inline_center(table)
        set_table_full_width(table)
        set_table_padding(table)

        # # --- B. 调用上面的终极边框函数 ---
        # # 这里设置：外框 1pt，内框 0.5pt，黑色
        set_table_borders_ultimate(table)

        # # --- C. 遍历行和单元格处理内容与行高 ---
        set_table_row_height(table)
        set_table_first_row_background(table)
        set_table_cell_align_line_height(table)

