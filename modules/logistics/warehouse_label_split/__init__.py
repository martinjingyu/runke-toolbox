"""入库标签 PDF 拆分：把一份汇总了很多箱子的入库标签 PDF（比如 CA1.pdf，每个仓库一份），
按「发货计划表里状态=未发货、发货时间=待定」的标签（SKU）拆成一个标签一个文件，见
splitter.py 顶部注释。
"""
from __future__ import annotations
