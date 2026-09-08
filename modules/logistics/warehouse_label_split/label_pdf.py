"""解析入库标签 PDF（比如 CA1.pdf）。

每个箱子占连续两页：
    第一页"Inbound"——收件人/仓库信息，没有 SKU 明细；
    第二页"Packing List"——这个箱子实际装的 SKU 明细（型号+数量），可能不止一行
                          （一箱混装多个 SKU）。
两页共享同一个 Box No.（形如 A8284826090800001.1），用它校验这两页确实是同一个箱子，
不是位置凑巧对上。

结构跟预期不一样（页数不是偶数、相邻两页 Box No. 对不上、找不到"Packing List"页、解析不出
SKU 明细）都不硬猜——直接报错，交给人工核对，跟 fba_label_redact/redact.py"结构不符合预期
就不动"是同一个原则。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

import fitz

_BOX_NO_RE = re.compile(r"Box No\.\s*\n(\S+)")


class LabelPdfStructureError(Exception):
    pass


@dataclass
class BoxItem:
    sku: str
    quantity: int


@dataclass
class Box:
    box_no: str
    inbound_page: int
    packing_page: int
    items: list[BoxItem]


def _box_no(text: str) -> str | None:
    m = _BOX_NO_RE.search(text)
    return m.group(1) if m else None


def _parse_packing_list_items(text: str) -> list[BoxItem]:
    """Packing List 页的正文是"No./BarCode/SKU/Qty"这四列表头，接着按行重复这四个字段——
    每一行渲染成纯文本后就是连续四行一组（编号、条码、型号、数量），一直到页面底部重复出现的
    "SKU:x  Qty:y"这行（跟 Inbound 页角落那行一样，是页面自带的水印，不是明细的一部分）为止。
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    try:
        start = lines.index("Qty") + 1
    except ValueError:
        raise LabelPdfStructureError("Packing List 页里找不到「Qty」表头")

    items: list[BoxItem] = []
    i = start
    while i + 3 < len(lines) and not lines[i].startswith("SKU:"):
        _no, _barcode, sku, qty = lines[i], lines[i + 1], lines[i + 2], lines[i + 3]
        if not qty.isdigit():
            raise LabelPdfStructureError(f"Packing List 页里「{sku}」这一行的数量「{qty}」不是数字，结构跟预期不符")
        items.append(BoxItem(sku=sku, quantity=int(qty)))
        i += 4
    if not items:
        raise LabelPdfStructureError("Packing List 页里没解析出任何 SKU 明细")
    return items


def parse_label_pdf(doc: fitz.Document) -> list[Box]:
    if doc.page_count % 2 != 0:
        raise LabelPdfStructureError(f"总页数是 {doc.page_count}，不是偶数，跟「每个箱子占两页」的预期结构不符")

    boxes: list[Box] = []
    for i in range(0, doc.page_count, 2):
        inbound_text = doc[i].get_text()
        packing_text = doc[i + 1].get_text()

        packing_lines = [line.strip() for line in packing_text.splitlines() if line.strip()]
        if not packing_lines or packing_lines[0] != "Packing List":
            raise LabelPdfStructureError(f"第 {i + 2} 页不是「Packing List」页，跟「每个箱子占两页」的预期结构不符")

        inbound_box_no = _box_no(inbound_text)
        packing_box_no = _box_no(packing_text)
        if inbound_box_no is None or packing_box_no is None:
            raise LabelPdfStructureError(f"第 {i + 1}/{i + 2} 页找不到「Box No.」")
        if inbound_box_no != packing_box_no:
            raise LabelPdfStructureError(
                f"第 {i + 1} 页 Box No.「{inbound_box_no}」和第 {i + 2} 页「{packing_box_no}」对不上"
            )

        items = _parse_packing_list_items(packing_text)
        boxes.append(Box(box_no=inbound_box_no, inbound_page=i, packing_page=i + 1, items=items))

    return boxes
