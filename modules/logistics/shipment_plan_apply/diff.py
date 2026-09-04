"""生成"变化前/变化后"预览用的数据快照——只挑这次操作真正碰到的那些行，不是整张表甩过来。

采购汇总表的行号在整个过程中不会变（这个表只插列，不插行），所以变化前/变化后可以直接按
同一个行号去读。发货计划汇总表会插行，行号会漂移。

同一个采购单号+型号在发货计划汇总表里可能同时有好几行"待定"（都是同一批还没决定去哪的库存，
见 shipment_summary.py 顶部说明），"变化前"就把这个采购单号+型号名下所有待定行都摆出来，
让人看到完整的库存状况；apply_shipment 可能会依次扣好几行、产生好几条 change，每条 change
最终都要能准确对应回它是从"变化前"哪一行扣出来的——不然预览里会把这一批操作产生的全部
新增行重复关联到每一条被删掉的原始行下面（同一采购单号+型号下有两行待定的话尤其明显）。

这里不是等 apply_plan 跑完之后把整张表从头扫一遍去"按数量猜"每条 change 对应哪一行——那样
一来行号会因为同一批里后面处理的记录插入位置更靠前而顶下去，二来这张表可能有几万行，扫描
次数还等于 change 数量，两者相乘在批量发货时非常慢（这也是这个工具偶尔跑到内存暴涨、卡死的
原因之一）。改成在 apply_plan 真正跑之前，先记下"行号 -> 这一行属于第几组"的映射，再跟着
apply_plan 实际发生的每一次 insert_rows 做等价的行号位移重放（apply_shipment 的规律是：
每一轮只消耗"当前还有余量的第一行"，一行一旦被 insert_above 部分扣减，这一轮就会立刻结束，
所以一行最多在一次 apply_shipment 调用里产生一条 change；这一行剩余的部分之后仍可能被后续
别的 change 继续消耗，即同一个分组可能对应好几条 change）——这样就能精确算出每条 change
最终落在哪一行、属于哪一组，不用扫表，复杂度只跟"这一批涉及多少条 change"有关，跟表格本身
有多少行无关。
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

    # row_of：活体行号 -> 这一行当前是什么身份，两种 token：
    #   ("grp", group_id)  这一行还是某个分组名下"待定"的那一行（可能是最初的，也可能是被
    #                       insert_above 扣过一部分之后留下的剩余待定行）
    #   ("ship", change_idx) 这一行是第 change_idx 条 change 产生的"已发货"记录
    # apply_plan 跑之前先按当前实际行号登记好所有分组的起点，再在下面跟着每条 change 实际发生
    # 的 insert_rows 做同样的位移，全程不用碰 worksheet，只是纯 Python 字典操作。
    row_of: dict[int, tuple] = {row[ROW_INDEX_KEY]: ("grp", row[GROUP_KEY]) for row in summary_before}

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

    def _shift_from(inserted_at: int) -> None:
        # insert_above 会在 inserted_at 这一行真的插入了一行，这一行（含）往下的所有已登记
        # 行号都要整体 +1，跟 worksheet 的实际位移保持一致——必须从大到小处理，不然会互相覆盖。
        for r in sorted((r for r in row_of if r >= inserted_at), reverse=True):
            row_of[r + 1] = row_of.pop(r)

    change_group_ids: list[int | None] = []
    for change in changes:
        src_row = change.new_row if change.kind == "insert_above" else change.pending_row
        src_token = row_of.get(src_row)
        group_id = src_token[1] if src_token is not None and src_token[0] == "grp" else None
        change_group_ids.append(group_id)

        if change.kind == "insert_above":
            # _shift_from 会顺带把刚查到的 ("grp", group_id) 从 new_row 搬到 new_row+1，
            # 正好等于 change.pending_row（剩余待定行的新位置），不用再手动搬一次。
            _shift_from(change.new_row)
            row_of[change.new_row] = ("ship", len(change_group_ids) - 1)
        else:
            row_of[change.pending_row] = ("ship", len(change_group_ids) - 1)

    final_row_of: dict[tuple, int] = {token: row for row, token in row_of.items()}

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
    for idx, (change, group_id) in enumerate(zip(changes, change_group_ids)):
        ship_row = final_row_of.get(("ship", idx))
        if ship_row is not None:
            row = _snapshot_summary(ship_row)
            row[GROUP_KEY] = group_id
            summary_after.append(row)
        if change.kind == "insert_above" and group_id is not None:
            # 剩余待定行如果之后又被别的 change 继续扣掉（同一分组名下不止一条 change），
            # 这个组号在 row_of 里早就被后面那条 change 的 ("ship", ...) 覆盖掉了，
            # final_row_of 这里自然查不到，不会重复显示一条已经不存在的"剩余待定行"。
            pending_row = final_row_of.get(("grp", group_id))
            if pending_row is not None:
                row = _snapshot_summary(pending_row)
                row[GROUP_KEY] = group_id
                summary_after.append(row)
        if progress_callback is not None:
            progress_callback("正在生成预览对比", idx + 1, total_changes)

    return PreviewResult(
        purchase=DiffTable(headers=purchase_headers, before_rows=purchase_before, after_rows=purchase_after),
        summary=DiffTable(headers=summary_headers, before_rows=summary_before, after_rows=summary_after),
        changes=changes,
    )
