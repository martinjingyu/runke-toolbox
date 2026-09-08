"""不同站点的入库标签 PDF，版式差别很大（CA1 是"一个箱子两页"，CG 是"一个箱子一页"），
但拆分/匹配/命名这套后续逻辑是通用的——用这几个数据结构把"这一页/这几页装的是什么"统一
成一样的形状，splitter.py 不用关心具体是哪个站点的版式。
"""
from __future__ import annotations

from dataclasses import dataclass


class LabelPdfStructureError(Exception):
    pass


@dataclass
class BoxItem:
    sku: str
    quantity: int


@dataclass
class Box:
    box_no: str
    pages: tuple[int, ...]  # 这个箱子占的页码，按顺序；CA1 是 (inbound_page, packing_page)，CG 是 (page,)
    items: list[BoxItem]
