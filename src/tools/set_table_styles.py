from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_table_borders(table, border_color="000000", border_size_pt=2):
    """
    设置表格边框样式
    :param table: docx 表格对象
    :param border_color: 16进制颜色字符串 (如 'FF0000' 为红色)
    :param border_size_pt: 边框宽度 (磅)，默认为 0.5pt
    """
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    # 获取或创建 tblBorders 节点
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is not None:
        tblPr.remove(tblBorders)

    tblBorders = OxmlElement('w:tblBorders')

    # 计算 Word 内部单位 (1 pt = 8 units)
    size_val = str(int(border_size_pt * 8))

    # 定义所有需要设置的边框位置
    # top, bottom, left, right, insideH (水平内线), insideV (垂直内线)
    borders_config = {
        'top': 'single',
        'bottom': 'single',
        'left': 'single',
        'right': 'single',
        'insideH': 'single',
        'insideV': 'single'
    }

    for tag, style in borders_config.items():
        element = OxmlElement(f'w:{tag}')
        element.set(qn('w:val'), style)      # 样式：single, double, dashed 等
        element.set(qn('w:sz'), size_val)     # 宽度
        element.set(qn('w:space'), '0')       # 间距
        element.set(qn('w:color'), border_color) # 颜色
        tblBorders.append(element)

    tblPr.append(tblBorders)

def set_cell_style(table):

    # --- 遍历行和单元格进行格式化 ---
    for i, row in enumerate(table.rows):
        # 设置行高自适应 (auto)
        tr_pr = row._tr.get_or_add_trPr()
        tr_height = tr_pr.find(qn('w:trHeight'))
        if tr_height is None:
            tr_height = OxmlElement('w:trHeight')
            tr_pr.append(tr_height)

        # hRule="auto" 表示高度随内容自动调整
        tr_height.set(qn('w:hRule'), 'auto')
        tr_height.set(qn('w:val'), '0')

        # 设置第一行背景颜色 (例如浅灰色 RGB: D9D9D9)
        if i == 0:
            for cell in row.cells:
                tc_pr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), 'D9D9D9') # 这里修改背景色 Hex 值
                tc_pr.append(shd)

        # 处理单元格内容和垂直居中
        for cell in row.cells:
            # 设置垂直居中
            tc_pr = cell._tc.get_or_add_tcPr()
            v_align = OxmlElement('w:vAlign')
            v_align.set(qn('w:val'), 'center')
            tc_pr.append(v_align)

            # 设置文字左对齐 & 字体大小 (可选)
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                # 统一字号
                for run in para.runs:
                    run.font.size = Pt(10.5) # 五号字


def set_table_styles(doc):
    for table in doc.tables:
        # --- 1. 获取或创建 tblPr (表格属性) ---
        # 修复报错的核心逻辑：手动查找或创建 <w:tblPr>
        tbl = table._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            # 将 tblPr 插入到 tbl 的最前面
            tbl.insert(0, tblPr)

        # --- 2. 设置表格宽度为 100% ---
        # 先删除旧的宽度设置，防止重复
        old_tblW = tblPr.find(qn('w:tblW'))
        if old_tblW is not None:
            tblPr.remove(old_tblW)

        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), '5000')   # 5000 代表 100% (Word单位是 1/50 %)
        tblW.set(qn('w:type'), 'pct')
        tblPr.append(tblW)

        # 设置表格对齐方式为居中（配合100%宽度使用效果最好）
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # --- 3. 设置单元格内边距 (Cell Margins) ---
        # --- 2. 设置单元格内边距 (Cell Margins) ---
        # 关键点：这里数值不能太大，否则会把格子撑得看起来像合并了
        # 单位是 twips (1/20 point)。这里设置为 54 (约 2.7pt)，比较适中
        tblCellMar = tblPr.find(qn('w:tblCellMar'))
        if tblCellMar is None:
            tblCellMar = OxmlElement('w:tblCellMar')
            tblPr.append(tblCellMar)

        margins = {'top': '54', 'bottom': '54', 'left': '54', 'right': '54'}
        for side, val in margins.items():
            tag = f'w:{side}'
            elem = tblCellMar.find(qn(tag))
            if elem is None:
                elem = OxmlElement(tag)
                tblCellMar.append(elem)
            elem.set(qn('w:w'), val)
            elem.set(qn('w:type'), 'dxa')


        set_cell_style(table)
        set_table_borders(table)



