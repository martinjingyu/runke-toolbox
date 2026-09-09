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
from .planner import Plan, apply_plan, apply_plan_purchase_only, apply_plan_summary_only
from .purchase_book import PurchaseBook
from .shipment_summary import ShipmentSummaryBook, ShipmentSummaryChange

__all__ = [
    "ROW_INDEX_KEY",
    "DiffTable",
    "PreviewResult",
    "run_and_capture_diff",
    "run_and_capture_diff_purchase_only",
    "run_and_capture_diff_summary_only",
]


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

    # 表头 -> 列号的映射只在这里算一次，不要让 read_row() 每读一行都自己重新扫一遍表头——
    # 真实表格 max_column 常年被撑得很大（openpyxl 的已知怪癖，见 column_utils.read_row 的
    # 说明），改动的行数一多，反复重扫表头的开销会成倍放大，是这个预览功能实测卡顿的主要
    # 原因。采购汇总表插入日期列之后列号会变，所以"改之前"和"改之后"要分别算一次；发货计划
    # 汇总表这边全程不会插列，一份映射从头用到尾就够。
    purchase_mapping = column_index_map(purchase_book.ws, purchase_book.header_row)
    summary_mapping = column_index_map(summary_book.ws, summary_book.header_row)

    def _snapshot_purchase(row_index: int, remaining: int, mapping: dict[str, int]) -> dict:
        # "未出货数量"是公式，直接读出来是公式文本不好看，预览里换成算好的数字
        row = read_row(purchase_book.ws, row_index, purchase_book.header_row, mapping)
        row["未出货数量"] = remaining
        row[ROW_INDEX_KEY] = row_index
        return row

    purchase_before = [
        _snapshot_purchase(r, rows_by_index[r].initial_remaining, purchase_mapping) for r in touched_purchase_rows
    ]

    def _snapshot_summary(row_index: int) -> dict:
        # 有些老数据"数量"这一格是公式（箱数*箱容），预览里换成算好的数字，不显示公式文本
        row = read_row(summary_book.ws, row_index, summary_book.header_row, summary_mapping)
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

    # 表头要在 apply_plan 跑完之后才重新取一遍——apply_plan 可能会往采购汇总表插一个新的
    # 日期列，插入前算的映射里不会有这一列（列号也可能因为插入整体右移），沿用旧的会导致
    # 新写进去的量从预览里彻底消失，或者读到错位的列。发货计划汇总表全程不插列，映射不用
    # 重算，沿用前面算好的那份。
    purchase_mapping = column_index_map(purchase_book.ws, purchase_book.header_row)
    purchase_headers = list(purchase_mapping.keys())
    summary_headers = list(summary_mapping.keys())

    purchase_after = [
        _snapshot_purchase(r, rows_by_index[r].remaining, purchase_mapping) for r in touched_purchase_rows
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

        # 不管拆分还是正好扣完，原来那条待定行（change.pending_row）最终都会变成已发货记录，
        # 直接读就是了——见 shipment_summary.py 类文档。
        row = _snapshot_summary(change.pending_row)
        row[GROUP_KEY] = group_id
        summary_after.append(row)

        if change.kind == "insert_new_pending":
            # 拆分：没扣完的部分被单独插了一行"待定"，固定在表格最下面，那才是这个分组接下来
            # 还剩多少待定库存——把分组指向这一行，不是刚变成已发货记录的 pending_row。这一批
            # 里如果后面还有分摊落到这个新插入的待定行上（change.pending_row 会等于这里的
            # new_row），group_by_pending_row 也要能查到同一个分组号，不然后续那条 change 会
            # 因为查不到分组、被当成孤立的一条处理。
            if group_id is not None:
                pending_row_by_group[group_id] = change.new_row
                group_by_pending_row[change.new_row] = group_id
        else:
            # convert_in_place：正好扣完，这个分组不再有待定库存了，从"还剩多少待定"里摘掉
            # （如果之前因为同一分组的前一条 change 记过，要撤掉，不然会重复展示一条早就
            # 不存在的"剩余待定行"）。
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


def run_and_capture_diff_purchase_only(
    plan: Plan,
    purchase_book: PurchaseBook,
    progress_callback=None,  # progress_callback(stage_label, done, total)
) -> DiffTable:
    """"采购订单分摊更新"用的精简版——只对比采购订单汇总表，完全不碰发货计划汇总表。
    跟 run_and_capture_diff 里"变化前/变化后"采购部分的写法是同一套（行号不会变，直接按
    行号读），只是没有发货计划汇总表那一半。
    """
    if plan.has_blocking_errors:
        raise ValueError("这一批发货计划里还有没解决的错误，不能写入")

    touched_rows = sorted({a.row.row_index for item in plan.items for a in item.allocations})
    rows_by_index = {r.row_index: r for r in purchase_book.rows}

    # 表头映射只算一次，不要让 read_row() 每读一行都重新扫一遍表头——见
    # column_utils.read_row 的说明，这是这个预览功能实测卡顿的主要原因。
    mapping = column_index_map(purchase_book.ws, purchase_book.header_row)

    def _snapshot(row_index: int, remaining: int, mapping: dict[str, int]) -> dict:
        row = read_row(purchase_book.ws, row_index, purchase_book.header_row, mapping)
        row["未出货数量"] = remaining
        row[ROW_INDEX_KEY] = row_index
        return row

    before_rows = [_snapshot(r, rows_by_index[r].initial_remaining, mapping) for r in touched_rows]

    apply_plan_purchase_only(
        plan,
        purchase_book,
        progress_callback=(
            (lambda done, total: progress_callback("正在写入变化", done, total))
            if progress_callback is not None
            else None
        ),
    )

    # 表头要在写完之后重新取一遍——可能往表格中间插了一个新的日期列，沿用写入前算的映射
    # 会导致新写进去的量从预览里彻底消失，或者读到错位的列。
    mapping = column_index_map(purchase_book.ws, purchase_book.header_row)
    headers = list(mapping.keys())
    after_rows = [_snapshot(r, rows_by_index[r].remaining, mapping) for r in touched_rows]

    return DiffTable(headers=headers, before_rows=before_rows, after_rows=after_rows)


def run_and_capture_diff_summary_only(
    plan: Plan,
    summary_book: ShipmentSummaryBook,
    progress_callback=None,  # progress_callback(stage_label, done, total)
) -> DiffTable:
    """"发货计划汇总表更新"用的精简版——只对比发货计划汇总表，完全不碰采购订单汇总表（那张
    表这里只读、不写，见 planner.py 的 apply_plan_summary_only）。跟 run_and_capture_diff
    里"变化前/变化后"发货计划部分的写法是同一套（待定行不会挪位置，直接按行号读；表头映射
    只算一次，不要让 read_row() 每读一行都重新扫一遍——见 column_utils.read_row 的说明，
    这是这个预览功能实测卡顿的主要原因），只是没有采购汇总表那一半。
    """
    if plan.has_blocking_errors:
        raise ValueError("这一批发货计划里还有没解决的错误，不能写入")

    mapping = column_index_map(summary_book.ws, summary_book.header_row)

    def _snapshot(row_index: int) -> dict:
        # 有些老数据"数量"这一格是公式（箱数*箱容），预览里换成算好的数字，不显示公式文本
        row = read_row(summary_book.ws, row_index, summary_book.header_row, mapping)
        row["数量"] = summary_book.read_quantity(row_index)
        row[ROW_INDEX_KEY] = row_index
        return row

    # 同一个采购单号+型号名下可能不止一条待定行（都是同一批还没决定去哪的库存），把这个
    # 采购单号+型号名下所有待定行都摆出来，让人看到完整的库存状况，不只是最后被扣的那一条。
    seen_keys: set = set()
    before_rows = []
    for item in plan.items:
        for allocation in item.allocations:
            key = (allocation.row.order_no, allocation.row.model)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            for r in summary_book.pending_rows(*key):
                row = _snapshot(r)
                row[GROUP_KEY] = len(before_rows)
                before_rows.append(row)

    # 待定行的行号 -> 它属于哪个分组。待定行不会挪位置，这张表从头到尾都有效。
    group_by_pending_row: dict[int, int] = {row[ROW_INDEX_KEY]: row[GROUP_KEY] for row in before_rows}

    changes = apply_plan_summary_only(
        plan,
        summary_book,
        progress_callback=(
            (lambda done, total: progress_callback("正在写入变化", done, total))
            if progress_callback is not None
            else None
        ),
    )
    # apply_shipment 是直接写、立刻生效的，change.new_row/pending_row 已经是真实、最终的
    # 行号，不用再额外模拟位移。

    headers = list(mapping.keys())

    after_rows = []
    total_changes = len(changes)
    # 分组名下"还剩多少待定"要展示最新状态，不能同一行出现好几遍——同一个分组名下如果这一批
    # 里连续被扣了好几次，只留最后一次的状态：还没扣完就记下它现在的行号，等所有 change 都
    # 处理完再统一取一次快照；扣完转正了就把这个分组摘掉，不再当"待定"展示。
    pending_row_by_group: dict[int, int] = {}
    for idx, change in enumerate(changes):
        group_id = group_by_pending_row.get(change.pending_row)

        # 不管拆分还是正好扣完，原来那条待定行（change.pending_row）最终都会变成已发货记录，
        # 直接读就是了——见 shipment_summary.py 类文档。
        row = _snapshot(change.pending_row)
        row[GROUP_KEY] = group_id
        after_rows.append(row)

        if change.kind == "insert_new_pending":
            # 拆分：没扣完的部分被单独插了一行"待定"，固定在表格最下面，那才是这个分组接下来
            # 还剩多少待定库存。这一批里如果后面还有分摊落到这个新插入的待定行上，
            # group_by_pending_row 也要能查到同一个分组号。
            if group_id is not None:
                pending_row_by_group[group_id] = change.new_row
                group_by_pending_row[change.new_row] = group_id
        else:
            # convert_in_place：正好扣完，这个分组不再有待定库存了，从"还剩多少待定"里摘掉。
            if group_id is not None:
                pending_row_by_group.pop(group_id, None)

        if progress_callback is not None:
            progress_callback("正在生成预览对比", idx + 1, total_changes)

    for group_id, pending_row in pending_row_by_group.items():
        row = _snapshot(pending_row)
        row[GROUP_KEY] = group_id
        after_rows.append(row)

    return DiffTable(headers=headers, before_rows=before_rows, after_rows=after_rows)
