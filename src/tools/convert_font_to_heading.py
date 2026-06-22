from docx import Document
from docx.shared import Pt
import copy


# 定义字号与样式的映射关系 (单位: pt)
FONT_SIZE_MAPPING = {
    17: ['Heading 1', '标题 1'],
    18: ['Heading 2', '标题 2'],  # 小二
    15: ['Heading 3', '标题 3'],  # 小三
}

def ensure_style_exists(doc, style_name):
    """
    检查文档中是否存在指定样式，如果不存在，则从标准模板中复制过来。
    """
    if style_name not in doc.styles:
        try:
            # 创建一个临时的标准文档作为“样式源”
            temp_doc = Document()
            source_style = temp_doc.styles[style_name]

            # 将样式XML节点深拷贝到当前文档的 styles 对象中
            # 注意：这是底层 XML 操作，能强制注入样式
            new_style_element = copy.deepcopy(source_style._element)
            doc.styles.element.append(new_style_element)
            print(f"[系统] 已自动注入缺失样式: {style_name}")
        except KeyError:
            print(f"[错误] 即使是标准模板中也找不到样式: {style_name}")

def convert_font_to_heading(doc):
    for para in doc.paragraphs:
        if not para.text.strip():
            continue

        # 获取段落字号
        font_size_pt = None
        for run in para.runs:
            if run.font.size is not None:
                font_size_pt = run.font.size.pt
                break

        if font_size_pt in FONT_SIZE_MAPPING:
            style_names = FONT_SIZE_MAPPING[font_size_pt]
            applied = False

            # 尝试列表中所有的样式名（先试英文，再试中文）
            for s_name in style_names:
                # 第一步：确保样式存在（解决警告的核心步骤）
                ensure_style_exists(doc, s_name)

                # 第二步：应用样式
                try:
                    para.style = doc.styles[s_name]
                    applied = True
                    break  # 应用成功就跳出循环
                except Exception as e:
                    continue

            if not applied:
                print(f"[警告] 无法为段落 '{para.text[:10]}...' 应用样式")

    print(f"处理完成")
