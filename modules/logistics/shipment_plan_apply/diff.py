"""生成"变化前/变化后"预览用的数据快照——只挑这次操作真正碰到的那些行，不是整张表甩过来。

采购汇总表的行号在整个过程中不会变（这个表只插列，不插行），所以变化前/变化后可以直接按
同一个行号去读。发货计划汇总表这边，待定行本身也不会挪位置（见 shipment_summary.py 顶部
说明——拆分只原地改数量，新插入的"已发货"记录固定放在表格最下面，不挨着待定行），所以这里
也可以直接按行号读，不用像插在待定行正上方那套做法一样，还要另外模拟一遍行号随插入位移的
过程。

同一个采购单号+型号在发货计划汇总表里可能同时有好几行"待定"（都是同一批还没决定去哪的库存，
见 shipment_summary.py 顶部说明），"变化前"就把这个采购单号+型号名下所有待定行都摆出来，
让人看到完整的库存状况；apply_shipment 可能会依次扣好几行、产生好几条 change，每条 change
最终都要能准确对应回它是从"变化前"哪一行扣出来的——不然预览里会把这一批操作产生的全部
新增行重复关联到每一条被删掉的原始行下面（同一采购单号+型号下有两行待定的话尤其明显）。
这里靠 change.pending_row（待定行原来的行号，扣完以后还是同一个）直接查表就行。
"""
from __future__ import annotations

from dataclasses import dataclass

from core.diff_preview import GROUP_KEY, ROW_INDEX_KEY, DiffTable

from .column_utils import column_index_map, read_row
from .planner import Plan, apply_plan
from .purchase_book import PurchaseBook
from .shipment_summary import ShipmentSummaryBook, ShipmentSummaryChange

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
    # 每一行顺手打一个组号（GROUP_KEY），后面精确匹配"变化后"就靠这个组号，不再靠采购单号+
    # 型号这种粗粒度的 key（见 core/diff_preview.py 里 GROUP_KEY 的说明）。
    seen_keys: set = set()
    summary_before = []
    for item in plan.items:
        for allocation in item.allocations:
            key = (allocation.row.order_no, allocation.row.model)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            for r in summary_book.pending_rows(*key):
                row = _snapshot_summary(r)
                row[GROUP_KEY] = len(summary_before)
                summary_before.append(row)

    # 待定行的行号 -> 它属于哪个分组。待定行不会挪位置（见 shipment_summary.py），这张表
    # 从头到尾都有效，不用像新行插在待定行正上方那套做法一样，还要跟着每一次插入模拟位移。
    group_by_pending_row: dict[int, int] = {row[ROW_INDEX_KEY]: row[GROUP_KEY] for row in summary_before}

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
    # apply_shipment 是直接写、立刻生效的（见 shipment_summary.py），change.new_row/
    # pending_row 已经是真实、最终的行号，不用再额外模拟位移。

    # 表头要在 apply_plan 跑完之后才取——apply_plan 可能会往采购汇总表插一个新的日期列，
    # 插入前取的表头列表里不会有这一列，写进去的量就会从预览里彻底消失（前后都不显示，
    # 而不是显示"从空到有值"这种正常的变化）
    purchase_headers = list(column_index_map(purchase_book.ws, purchase_book.header_row).keys())
    summary_headers = list(column_index_map(summary_book.ws, summary_book.header_row).keys())

    purchase_after = [
        _snapshot_purchase(r, rows_by_index[r].remaining) for r in touched_purchase_rows
    ]

    summary_after = []
    total_changes = len(changes)
    # 分组名下"还剩多少待定"要展示最新状态，不能同一行出现好几遍——同一个分组名下如果这一批
    # 里连续被扣了好几次（先拆一部分，后面又把剩下的扣完），只留最后一次的状态：还没扣完就
    # 记下它现在的行号，等所有 change 都处理完再统一取一次快照；扣完转正了就把这个分组摘掉，
    # 不再当"待定"展示。
    pending_row_by_group: dict[int, int] = {}
    for idx, change in enumerate(changes):
        group_id = group_by_pending_row.get(change.pending_row)

        if change.kind == "insert_above":
            # 新插入的"已发货"记录固定在表格最下面，跟这一行原来在表格哪个位置没关系，
            # 直接读就是了。
            row = _snapshot_summary(change.new_row)
            row[GROUP_KEY] = group_id
            summary_after.append(row)
            if group_id is not None:
                pending_row_by_group[group_id] = change.pending_row
        else:
            # convert_in_place：待定行原地转正，不再是待定库存了，从"还剩多少待定"里摘掉
            # （如果之前因为同一分组的前一条 change 记过，要撤掉，不然会重复展示一条早就
            # 不存在的"剩余待定行"）。
            row = _snapshot_summary(change.pending_row)
            row[GROUP_KEY] = group_id
            summary_after.append(row)
            if group_id is not None:
                pending_row_by_group.pop(group_id, None)

        if progress_callback is not None:
            progress_callback("正在生成预览对比", idx + 1, total_changes)

    for group_id, pending_row in pending_row_by_group.items():
        row = _snapshot_summary(pending_row)
        row[GROUP_KEY] = group_id
        summary_after.append(row)

    return PreviewResult(
        purchase=DiffTable(headers=purchase_headers, before_rows=purchase_before, after_rows=purchase_after),
        summary=DiffTable(headers=summary_headers, before_rows=summary_before, after_rows=summary_after),
        changes=changes,
    )
