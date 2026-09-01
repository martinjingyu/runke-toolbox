"""发货数量核对（Walmart）：拆分箱唛 PDF + 核对实际发货数量 vs 发货计划表。

三个输入：箱唛 PDF、翻译表（WM-SKU→货号）、发货计划表。详细的匹配链路见 reconcile.py 顶部注释。

注意：这个包的 __init__.py 故意不在这里 import reconcile/box_label 这些重量级模块（会带出
fitz/pylibdmtx/pytesseract）——"物流仓库"部门的工具列表页只需要 tesseract_dependency.py
（纯标准库）就能列出这个工具、检查依赖，不该因为列了这一项就先把这些库都导入一遍。
真正用到 run()/write_report_xlsx() 的地方直接 `from .reconcile import ...`。
"""
