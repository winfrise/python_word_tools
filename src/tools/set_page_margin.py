from docx import Document
from docx.shared import Cm
from docx.oxml.ns import qn

def set_page_margin(doc):
    #  遍历文档中的所有节（sections）
    for section in doc.sections:
        section.top_margin = Cm(2.54)      # 上边距
        section.bottom_margin = Cm(2.54)   # 下边距
        section.left_margin = Cm(1.9)     # 左边距
        section.right_margin = Cm(1.9)    # 右边距


        # 页面级别，禁用对齐到网格
        sectPr = section._sectPr
        docGrid = sectPr.find(qn('w:docGrid'))
        if docGrid is not None:
            sectPr.remove(docGrid)

    print('设置页面边距')

    return doc
