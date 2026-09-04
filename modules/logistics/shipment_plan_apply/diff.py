"""生成"变化前/变化后"预览用的数据快照——只挑这次操作真正碰到的那些行，不是整张表甩过来。

采购汇总表的行号在整个过程中不会变（这个表只插列，不插行），所以变化前/变化后可以直接按
同一个行号去读。发货计划汇总表会插行，行号会漂移，所以变化后不能信"处理过程中记下来的行号"
（同一批里后面处理的记录插入位置更靠前的话，会把前面已经记下的行号顶下去）——统一改成按
内容重新查一遍。

同一个采购单号+型号在发货计划汇总表里可能同时有好几行"待定"（都是同一批还没决定去哪的库存，
见 shipment_summary.py 顶部说明），"变化前"就把这个采购单号+型号名下所有待定行都摆出来，
让人看到完整的库存状况；"变化后"按每条 change 具体的量去精确匹配（发出的那行数量要等于
change.quantity，插入型的话剩余待定行数量要等于 change.pending_remaining_after），因为
apply_shipment 可能会依次扣好几行、产生好几条 change。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass

from core.diff_preview import ROW_INDEX_KEY, DiffTable

from .column_utils import column_index_map, read_row
from .planner import Plan, apply_plan
from .purchase_book import PurchaseBook
from .shipment_summary import PENDING_LABEL, ShipmentSummaryBook, ShipmentSummaryChange

__all__ = ["ROW_INDEX_KEY", "DiffTable", "PreviewResult", "run_and_capture_diff"]


@dataclass
class PreviewResult:
    purchase: DiffTable
    summary: DiffTable
    changes: list[ShipmentSummaryChange]


def run_and_capture_diff(
    plan: Plan,
    purchase_book: PurchaseBook,
    summary_book: ShipmentSummaryBook,
    progress_callback=None,  # progress_callback(stage_label, done, total)
) -> PreviewResult:
    if plan.has_blocking_errors:
        raise ValueError("这一批发货计划里还有没解决的错误，不能写入")

    touched_purchase_rows = sorted({a.row.row_index for item in plan.items for a in item.allocations})
    rows_by_index = {r.row_index: r for r in purchase_book.rows}

    def _snapshot_purchase(row_index: int, remaining: int) -> dict:
        # "未出货数量"是公式，直接读出来是公式文本不好看，预览里换成算好的数字
        row = read_row(purchase_book.ws, row_index, purchase_book.header_row)
        row["未出货数量"] = remaining
        row[ROW_INDEX_KEY] = row_index
        return row

    purchase_before = [
        _snapshot_purchase(r, rows_by_index[r].initial_remaining) for r in touched_purchase_rows
    ]

    def _snapshot_summary(row_index: int) -> dict:
        # 有些老数据"数量"这一格是公式（箱数*箱容），预览里换成算好的数字，不显示公式文本
        row = read_row(summary_book.ws, row_index, summary_book.header_row)
        row["数量"] = summary_book.read_quantity(row_index)
        row[ROW_INDEX_KEY] = row_index
        return row

    # 同一个采购单号+型号名下可能不止一条待定行（都是同一批还没决定去哪的库存），把这个
    # 采购单号+型号名下所有待定行都摆出来，让人看到完整的库存状况，不只是最后被扣的那一条。
    seen_keys: set = set()
    summary_before = []
    for item in plan.items:
        for allocation in item.allocations:
            key = (allocation.row.order_no, allocation.row.model)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            for r in summary_book.pending_rows(*key):
                summary_before.append(_snapshot_summary(r))

    changes = apply_plan(
        plan,
        purchase_book,
        summary_book,
        progress_callback=(
            (lambda done, total: progress_callback("正在写入变化", done, total))
            if progress_callback is not None
            else None
        ),
    )

    # 表头要在 apply_plan 跑完之后才取——apply_plan 可能会往采购汇总表插一个新的日期列，
    # 插入前取的表头列表里不会有这一列，写进去的量就会从预览里彻底消失（前后都不显示，
    # 而不是显示"从空到有值"这种正常的变化）
    purchase_headers = list(column_index_map(purchase_book.ws, purchase_book.header_row).keys())
    summary_headers = list(column_index_map(summary_book.ws, summary_book.header_row).keys())

    purchase_after = [
        _snapshot_purchase(r, rows_by_index[r].remaining) for r in touched_purchase_rows
    ]

    # 变化后同理，不能只按采购单号+型号广撒网——按每条 change 具体的量去精确匹配（发出的那行
    # 数量要等于 change.quantity，插入型的话剩余待定行数量要等于 change.pending_remaining_after），
    # seen_rows 防止两条 change 抢到同一行。
    summary_after = []
    seen_after_rows: set = set()
    c_order = summary_book.col["采购单号"]
    c_model = summary_book.col["型号"]
    c_ship = summary_book.col["发货时间"]
    total_changes = len(changes)
    for done, change in enumerate(changes, start=1):
        for r in range(summary_book.header_row + 1, summary_book.ws.max_row + 1):
            if r in seen_after_rows:
                continue
            if (
                summary_book.ws.cell(row=r, column=c_order).value != change.order_no
                or summary_book.ws.cell(row=r, column=c_model).value != change.model
            ):
                continue
            ship_val = summary_book.ws.cell(row=r, column=c_ship).value
            is_this_shipment = isinstance(ship_val, dt.datetime) and ship_val.date() == plan.ship_date
            if is_this_shipment and summary_book.read_quantity(r) == change.quantity:
                seen_after_rows.add(r)
                summary_after.append(_snapshot_summary(r))
                continue
            if (
                change.kind == "insert_above"
                and ship_val == PENDING_LABEL
                and summary_book.read_quantity(r) == change.pending_remaining_after
            ):
                seen_after_rows.add(r)
                summary_after.append(_snapshot_summary(r))
        if progress_callback is not None:
            progress_callback("正在生成预览对比", done, total_changes)

    return PreviewResult(
        purchase=DiffTable(headers=purchase_headers, before_rows=purchase_before, after_rows=purchase_after),
        summary=DiffTable(headers=summary_headers, before_rows=summary_before, after_rows=summary_after),
        changes=changes,
    )
