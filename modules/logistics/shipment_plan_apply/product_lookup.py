"""在售产品信息总表：把运营发货计划表里的 SKU（AMZ-SKU 或 RK-SKU）翻译成货号（"同款"字段），
查不到就是硬错误——宁可报错也不能猜一个可能不对的货号，这会导致真实库存记错到别的产品上。

"变体"字段不用于这里的库存匹配——核对过真实数据后发现它标的是"同一个造型、不同颜色"的一组
货号（比如水滴白色/灰色/亮棕色树脂台灯共用一个变体标签），颜色不一样的货没法互相顶替发货。
一个货号自己名下的采购订单没有足够库存时，就是真的缺货，直接报错让人处理，不要找别的货号
的库存来凑。
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import openpyxl

from .column_utils import column_index_map, find_header_row, require_columns

REQUIRED_HEADERS = ["AMZ-SKU", "RK-SKU", "同款"]


class ProductLookupError(Exception):
    pass


@dataclass
class ProductLookup:
    amz_to_huohao: dict[str, str]
    rk_to_huohao: dict[str, str]

    def resolve(self, sku_kind: str, sku: str) -> str | None:
        """先按 RK-SKU 查，查不到再按 AMZ-SKU 查（不管 sku_kind 是哪种——运营表里填错列、
        或者同一个值本来就两边都用的情况都能兜住）。sku_kind 保留是为了兼容调用方/历史签名，
        不参与判断了。两边都查不到返回 None（由调用方决定怎么报错）。"""
        huohao = self.rk_to_huohao.get(sku)
        if huohao is not None:
            return huohao
        return self.amz_to_huohao.get(sku)


def load_product_lookup(path: Path) -> ProductLookup:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    header_row = find_header_row(ws, REQUIRED_HEADERS)
    cols = column_index_map(ws, header_row)
    idx = require_columns(cols, REQUIRED_HEADERS, "在售产品信息总表")
    amz_col = idx["AMZ-SKU"]
    rk_col = idx["RK-SKU"]
    huohao_col = idx["同款"]

    amz_to_huohao: dict[str, str] = {}
    rk_to_huohao: dict[str, str] = {}
    conflicts: list[str] = []

    for row in ws.iter_rows(min_row=header_row + 1):
        amz = row[amz_col - 1].value
        rk = row[rk_col - 1].value
        huohao = row[huohao_col - 1].value

        if huohao is None:
            continue
        huohao = str(huohao).strip()
        if not huohao:
            continue

        if amz is not None:
            amz = str(amz).strip()
            if amz:
                conflicts.extend(_record(amz_to_huohao, amz, huohao, f"AMZ-SKU「{amz}」"))

        if rk is not None:
            rk = str(rk).strip()
            if rk:
                conflicts.extend(_record(rk_to_huohao, rk, huohao, f"RK-SKU「{rk}」"))

    if conflicts:
        raise ProductLookupError(
            "在售产品信息总表里有数据冲突，同一个值映射到了不止一个货号，"
            "无法确定该用哪一个，需要人工先核对表格：\n" + "\n".join(conflicts)
        )

    return ProductLookup(amz_to_huohao=amz_to_huohao, rk_to_huohao=rk_to_huohao)


def _record(table: dict[str, str], key: str, value: str, label: str) -> list[str]:
    existing = table.get(key)
    if existing is None:
        table[key] = value
        return []
    if existing != value:
        return [f"{label}：同时映射到了「{existing}」和「{value}」"]
    return []
