"""解析 CG 站点的入库标签 PDF（比如 CG-TS-MD.pdf）。

版式比 CA1 简单：每一页就是一个箱子的标签，没有"Inbound/Packing List 两页一组"的区分——
这一页上「PART#:」后面那个值就是这个箱子的标签（SKU），「Carton Quantity:」后面是这一箱的
数量。不需要靠箱子编号去重（每页天然对应一个箱子，不会像 CA1 那样一箱混装多个 SKU、需要用
Box No. 把同一个箱子的多行明细归到一起），直接拿页码本身当箱子标识就够用。

一页里找不到「PART#:」或「Carton Quantity:」都不硬猜——报错交给人工核对，跟 CA1 的
label_pdf.py 是同一个原则。
"""
from __future__ import annotations

import re

import fitz

from .box import Box, BoxItem, LabelPdfStructureError

_PART_RE = re.compile(r"PART#:\s*(\S+)")
_QTY_RE = re.compile(r"Carton Quantity:\s*(\d+)")


def parse_label_pdf(doc: fitz.Document) -> list[Box]:
    boxes: list[Box] = []
    for i in range(doc.page_count):
        text = doc[i].get_text()

        part_match = _PART_RE.search(text)
        if part_match is None:
            raise LabelPdfStructureError(f"第 {i + 1} 页找不到「PART#:」，结构跟预期不符")

        qty_match = _QTY_RE.search(text)
        if qty_match is None:
            raise LabelPdfStructureError(f"第 {i + 1} 页找不到「Carton Quantity:」，结构跟预期不符")

        sku = part_match.group(1).strip()
        quantity = int(qty_match.group(1))
        boxes.append(Box(box_no=f"page-{i}", pages=(i,), items=[BoxItem(sku=sku, quantity=quantity)]))

    return boxes
