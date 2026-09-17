"""把 fba_source 解析出来的"MSKU -> FBA 号"结果，跟发货计划表（"发货计划(2)"这张表，跟
shipment_plan_apply 模块里"发货计划汇总表"用的是同一套列名）匹配起来，按箱数拆分对应的行。

跟 shipment_plan_apply/planner.py 一样分两段："算一遍"（build_plan，只读，不管有没有问题
都会跑完、把所有问题一次性收集出来）和"真的写"（apply_plan，只有 build_plan 判定完全没有
问题才能调用）。不一样的地方在于插入位置——这次业务上要求拆分出来的行必须留在原来那一行
附近，不能像 shipment_summary.py 那样统一插到表格最下面，所以 apply_plan 用的是"整表按行
号从大到小做一次性搬移"，不逐次调用 openpyxl 的 insert_rows()（它每次插入不管插入点在哪，
内部都要把整张表的单元格坐标全部重排一次，逐次调用会很慢），细节见 _apply_row_shift 的
说明。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from openpyxl.worksheet.worksheet import Worksheet

from modules.logistics.shipment_plan_apply.column_utils import (
    column_index_map,
    copy_row,
    find_header_row,
    require_columns,
    resolve_cell_value,
    unmerge_overlapping_rows,
)

from .fba_source import FbaEntry, FbaSourceResult

REQUIRED_HEADERS = ["采购单号", "标签", "箱数", "箱容", "仓库", "FBA ID", "追踪编号", "ZD", "发货时间"]

_VALID_ZD = {"US", "CA"}


def _to_date(value) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return None


@dataclass
class PlanRow:
    row_index: int
    label: str
    order_no: str
    boxes: float
    box_capacity: float


@dataclass
class SplitPiece:
    boxes: float
    fba_id: str
    destination: str
    internal_id: str


@dataclass
class RowPlan:
    """一个原始行最终要变成的样子——只有一段就是原地改几个字段，好几段就是原行+新增行，
    新增的行都紧跟在这一行原来的位置后面。
    """

    row_index: int
    pieces: list[SplitPiece]


@dataclass
class Plan:
    ship_date: dt.date
    row_plans: list[RowPlan]
    errors: list[str]
    skipped_fba_rows: list = field(default_factory=list)
    duplicates_removed: int = 0
    header_row: int = 0
    col: dict[str, int] = field(default_factory=dict)
    original_last_row: int = 0

    @property
    def has_blocking_errors(self) -> bool:
        return bool(self.errors)


def build_plan(ws: Worksheet, fba_result: FbaSourceResult, ship_date: dt.date) -> Plan:
    header_row = find_header_row(ws, REQUIRED_HEADERS, max_scan_rows=10, context="发货计划表")
    cols = column_index_map(ws, header_row)
    col = require_columns(cols, REQUIRED_HEADERS, "发货计划表")

    relevant_labels = set(fba_result.entries_by_msku)
    groups: dict[str, list[PlanRow]] = {}
    bad_labels: set[str] = set()
    errors: list[str] = []

    last_row = ws.max_row
    for r in range(header_row + 1, last_row + 1):
        label_value = resolve_cell_value(ws, r, col["标签"])
        if label_value is None:
            continue
        label = str(label_value).strip()
        if not label or label not in relevant_labels:
            continue

        row_date = _to_date(ws.cell(row=r, column=col["发货时间"]).value)
        if row_date != ship_date:
            continue

        boxes_value = ws.cell(row=r, column=col["箱数"]).value
        capacity_value = ws.cell(row=r, column=col["箱容"]).value
        row_ok = True
        if not isinstance(boxes_value, (int, float)) or isinstance(boxes_value, bool):
            errors.append(f"第{r}行（标签「{label}」）：箱数不是数字，读不出来")
            bad_labels.add(label)
            row_ok = False
        if not isinstance(capacity_value, (int, float)) or isinstance(capacity_value, bool):
            errors.append(f"第{r}行（标签「{label}」）：箱容不是数字，读不出来")
            bad_labels.add(label)
            row_ok = False
        if not row_ok:
            continue

        fba_id_value = ws.cell(row=r, column=col["FBA ID"]).value
        if fba_id_value not in (None, ""):
            errors.append(
                f"第{r}行（标签「{label}」）：已经有 FBA ID「{fba_id_value}」，这种情况不该出现——"
                "这一行不应该落在这次待处理范围里，需要人工核实"
            )
            bad_labels.add(label)

        zd_value = ws.cell(row=r, column=col["ZD"]).value
        zd = str(zd_value).strip() if zd_value is not None else ""
        if zd not in _VALID_ZD:
            errors.append(f"第{r}行（标签「{label}」）：ZD 不是预期的 US/CA（实际是「{zd_value}」），没法拼出仓库")
            bad_labels.add(label)

        order_no_value = ws.cell(row=r, column=col["采购单号"]).value
        groups.setdefault(label, []).append(
            PlanRow(
                row_index=r,
                label=label,
                order_no=str(order_no_value) if order_no_value is not None else "",
                boxes=boxes_value,
                box_capacity=capacity_value,
            )
        )

    for label in sorted(relevant_labels - groups.keys()):
        errors.append(f"MSKU「{label}」在 FBA 数据里有记录，但发货计划表里没有找到发货日期={ship_date} 且标签匹配的行")

    row_plans: list[RowPlan] = []
    for label, plan_rows in groups.items():
        if label in bad_labels:
            continue

        capacities = {pr.box_capacity for pr in plan_rows}
        if len(capacities) > 1:
            errors.append(f"标签「{label}」这一批里箱容不一致：{sorted(capacities)}，没法换算箱数")
            continue
        box_capacity = plan_rows[0].box_capacity

        entries = fba_result.entries_by_msku[label]
        converted_entries: list[tuple[FbaEntry, int]] = []
        conversion_failed = False
        for entry in entries:
            if box_capacity <= 0 or entry.quantity_pieces % box_capacity != 0:
                errors.append(
                    f"标签「{label}」的 FBA 号「{entry.fba_id}」预计商品数量 {entry.quantity_pieces} "
                    f"除不尽箱容 {box_capacity}，没法换算成箱数"
                )
                conversion_failed = True
                continue
            converted_entries.append((entry, int(entry.quantity_pieces / box_capacity)))
        if conversion_failed:
            continue

        total_plan_boxes = sum(pr.boxes for pr in plan_rows)
        total_fba_boxes = sum(boxes for _, boxes in converted_entries)
        if total_plan_boxes != total_fba_boxes:
            errors.append(
                f"标签「{label}」箱数对不上：发货计划表这几行加起来是 {total_plan_boxes} 箱，"
                f"FBA 数据换算出来是 {total_fba_boxes} 箱"
            )
            continue

        row_plans.extend(_zip_split(plan_rows, converted_entries))

    if errors:
        row_plans = []

    return Plan(
        ship_date=ship_date,
        row_plans=row_plans,
        errors=errors,
        skipped_fba_rows=fba_result.skipped,
        duplicates_removed=fba_result.duplicates_removed,
        header_row=header_row,
        col=col,
        original_last_row=last_row,
    )


def _zip_split(plan_rows: list[PlanRow], converted_entries: list[tuple[FbaEntry, int]]) -> list[RowPlan]:
    """核心拆分算法：按顺序"拉链式"消耗两个数量序列——计划表这几行的箱数（供给）跟 FBA 号
    换算出来的箱数（需求）。每一步都切一刀 min(当前行剩余箱数, 当前 FBA 号剩余箱数)，切到
    哪个先归零就换下一个。调用方已经校验过两边总数相等，正常不会出现某一边提前用完的情况，
    真出现了说明前面的校验漏了什么，直接报内部错误，不猜、不硬凑。
    """
    row_plans: list[RowPlan] = []
    row_iter = iter(plan_rows)
    entry_iter = iter(converted_entries)

    def next_row():
        pr = next(row_iter, None)
        return pr, (pr.boxes if pr is not None else 0)

    def next_entry():
        nxt = next(entry_iter, None)
        if nxt is None:
            return None, 0
        return nxt[0], nxt[1]

    current_row, row_remaining = next_row()
    current_entry, entry_remaining = next_entry()
    pieces: list[SplitPiece] = []

    while current_row is not None:
        if current_entry is None:
            raise AssertionError(
                "拆分算法内部错误：FBA 号提前用完，但计划行还没处理完——"
                "箱数总和校验应该已经挡住了这种情况，不应该走到这里"
            )
        take = min(row_remaining, entry_remaining)
        pieces.append(
            SplitPiece(
                boxes=take,
                fba_id=current_entry.fba_id,
                destination=current_entry.destination,
                internal_id=current_entry.internal_id,
            )
        )
        row_remaining -= take
        entry_remaining -= take

        if row_remaining == 0:
            row_plans.append(RowPlan(row_index=current_row.row_index, pieces=pieces))
            pieces = []
            current_row, row_remaining = next_row()
        if entry_remaining == 0:
            current_entry, entry_remaining = next_entry()

    return row_plans


def apply_plan(ws: Worksheet, plan: Plan) -> None:
    if plan.has_blocking_errors:
        raise ValueError("这一批还有没解决的问题，不能写入")
    if not plan.row_plans:
        return

    col = plan.col
    header_row = plan.header_row
    last_row = plan.original_last_row
    max_col = ws.max_column
    plans_by_row = {rp.row_index: rp for rp in plan.row_plans}

    # 前缀和：new_start[r] = r 加上"r 之前所有目标行因为拆分多出来的行数"之和。这个映射
    # 严格单调递增，不同原始行对应的新行区间互不重叠，见模块顶部说明。
    new_start: dict[int, int] = {}
    offset = 0
    for r in range(header_row + 1, last_row + 1):
        new_start[r] = r + offset
        k = len(plans_by_row[r].pieces) if r in plans_by_row else 1
        offset += k - 1
    final_last_row = last_row + offset

    if offset > 0:
        unmerge_overlapping_rows(ws, header_row + 1, final_last_row)

    for r in range(last_row, header_row, -1):
        dest_start = new_start[r]
        row_plan = plans_by_row.get(r)

        if row_plan is None:
            if dest_start != r:
                copy_row(ws, dest_start, r, max_col=max_col)
            continue

        pieces = row_plan.pieces
        # 先把所有需要新增的行从原行 r 复制出来（这时候 r 本身的字段还是原样，没被覆盖过），
        # 再统一去写各段的箱数/FBA ID/仓库/追踪编号——顺序不能反，不然后面复制的行会抄到
        # 已经被前一段覆盖过的字段，见模块顶部说明。
        for i in range(1, len(pieces)):
            copy_row(ws, dest_start + i, r, max_col=max_col)
        if dest_start != r:
            copy_row(ws, dest_start, r, max_col=max_col)

        for i, piece in enumerate(pieces):
            _write_piece_fields(ws, dest_start + i, col, piece)

    _sync_auto_filter(ws, header_row, max_col, final_last_row)


def _write_piece_fields(ws: Worksheet, row: int, col: dict[str, int], piece: SplitPiece) -> None:
    ws.cell(row=row, column=col["箱数"]).value = piece.boxes
    ws.cell(row=row, column=col["FBA ID"]).value = piece.fba_id
    ws.cell(row=row, column=col["追踪编号"]).value = piece.internal_id
    zd_value = ws.cell(row=row, column=col["ZD"]).value
    zd = str(zd_value).strip() if zd_value is not None else ""
    ws.cell(row=row, column=col["仓库"]).value = f"{zd}({piece.destination})"


def _sync_auto_filter(ws: Worksheet, header_row: int, max_col: int, final_last_row: int) -> None:
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    current_ref = ws.auto_filter.ref
    if not current_ref:
        return
    min_col, min_row, max_col_existing, max_row_existing = range_boundaries(current_ref)
    target_max_row = max(max_row_existing, final_last_row)
    target_max_col = max(max_col_existing, max_col)
    ws.auto_filter.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(target_max_col)}{target_max_row}"
