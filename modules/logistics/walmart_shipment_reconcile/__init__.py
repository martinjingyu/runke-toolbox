"""发货数量核对（Walmart）：拆分箱唛 PDF + 核对实际发货数量 vs 发货计划表。

三个输入：箱唛 PDF、翻译表（WM-SKU→货号）、发货计划表。详细的匹配链路见 reconcile.py 顶部注释。
"""
from .reconcile import RunReport, run, write_report_xlsx

__all__ = ["run", "write_report_xlsx", "RunReport"]
