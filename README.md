# 基于Python + python-docx 实现的word排版工具

```
# 1. 创建虚拟环境
python -m venv venv

# 2. 激活环境 (Windows)
source venv/bin/activate

# 3. 安装 python-docx
pip install python-docx

# 导出依赖列表
pip freeze > requirements.txt
```

注意事项：

一、 行距问题

1. 检查是否开启网格。页面 -> 页面设置 -> 文档网络（不要选择无网络）
2. 全选页面中的内容 -> 开始 -> 段落 -> 缩进和间距 ->  间距 -> 【勾选】如果定义了文档网格，则与网格对齐