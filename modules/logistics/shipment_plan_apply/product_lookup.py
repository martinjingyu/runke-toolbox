"""在售产品信息总表：把运营发货计划表里的 SKU（AMZ-SKU 或 RK-SKU）翻译成货号（"同款"字段），
查不到就是硬错误——宁可报错也不能猜一个可能不对的货号，这会导致真实库存记错到别的产品上。

"变体"字段是兜底用的：有些货物本身是同一个东西，只是被打了不同的货号在卖，"变体"记录的就是
"这个货号背后真正用来对账/采购的那个货号"。货号在采购汇总表里找不到剩余库存时，才会用这个
字段换一个货号再查一次（见 purchase_allocation.py）。
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import openpyxl

from .column_utils import column_index_map, find_header_row, require_columns

REQUIRED_HEADERS = ["AMZ-SKU", "RK-SKU", "同款", "变体"]


class ProductLookupError(Exception):
    pass


@dataclass
class ProductLookup:
    amz_to_huohao: dict[str, str]
    rk_to_huohao: dict[str, str]
    huohao_to_variant: dict[str, str]

    def resolve(self, sku_kind: str, sku: str) -> str | None:
        """sku_kind 是 "AMZ" 或 "RK"。查不到返回 None（由调用方决定怎么报错）。"""
        table = self.amz_to_huohao if sku_kind == "AMZ" else self.rk_to_huohao
        return table.get(sku)

    def variant_of(self, huohao: str) -> str | None:
        variant = self.huohao_to_variant.get(huohao)
        if variant is None or variant == huohao:
            return None
        return variant


def load_product_lookup(path: Path) -> ProductLookup:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    header_row = find_header_row(ws, REQUIRED_HEADERS)
    cols = column_index_map(ws, header_row)
    idx = require_columns(cols, REQUIRED_HEADERS, "在售产品信息总表")
    amz_col = idx["AMZ-SKU"]
    rk_col = idx["RK-SKU"]
    huohao_col = idx["同款"]
    variant_col = idx["变体"]

    amz_to_huohao: dict[str, str] = {}
    rk_to_huohao: dict[str, str] = {}
    huohao_to_variant: dict[str, str] = {}
    conflicts: list[str] = []

    for row in ws.iter_rows(min_row=header_row + 1):
        amz = row[amz_col - 1].value
        rk = row[rk_col - 1].value
        huohao = row[huohao_col - 1].value
        variant = row[variant_col - 1].value

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

        if variant is not None:
            variant = str(variant).strip()
            if variant:
                conflicts.extend(
                    _record(huohao_to_variant, huohao, variant, f"货号「{huohao}」的变体")
                )

    if conflicts:
        raise ProductLookupError(
            "在售产品信息总表里有数据冲突，同一个值映射到了不止一个货号/变体，"
            "无法确定该用哪一个，需要人工先核对表格：\n" + "\n".join(conflicts)
        )

    return ProductLookup(
        amz_to_huohao=amz_to_huohao,
        rk_to_huohao=rk_to_huohao,
        huohao_to_variant=huohao_to_variant,
    )


def _record(table: dict[str, str], key: str, value: str, label: str) -> list[str]:
    existing = table.get(key)
    if existing is None:
        table[key] = value
        return []
    if existing != value:
        return [f"{label}：同时映射到了「{existing}」和「{value}」"]
    return []
