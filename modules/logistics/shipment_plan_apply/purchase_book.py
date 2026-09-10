"""采购订单汇总表：每一行是"一笔采购订单里的一个型号"，后面跟着一长串"日期列"（哪一批次
出货，就在那一列写发出的数量），"未出货数量"是表格里已经有的本地公式：
    =+<订单数量列><行>-SUM(<第一个日期列><行>:<最后一个日期列><行>)
本工具只需要把这次发货的数量填进正确的日期列——公式本身在 Excel 打开时会自己算出新的
"未出货数量"，不用我们手动改那一格。

**这张表不再自动插入日期列**（早期版本会在目标日期没有对应列时现场插入一新列，但插入会
挪动这一列右边所有列的位置，即便小心处理了公式区间/格式/列宽，实际使用中还是出现过插入点
右边的公式跟着错位的问题）。现在的规则是：日期列必须已经存在于表格里，才能往里面写数据——
不存在就直接报错（`DateColumnNotFoundError`），提示人工先在表格里加好这一天的列，不再由
代码插入。`find_date_column` 只做只读查找，找不到返回 None；`require_date_column` 在此基础
上找不到就报错，是"采购订单分摊更新"/"发货计划自动更新"实际写入前要用的那个。

分摊规则：同一个货号（型号）可能对应好几笔采购订单，按采购日期从早到晚，先把早的写满，
写满了（未出货数量到 0）就换下一笔。如果找不到货号对应的订单，或者所有能找到的订单加起来
的未出货数量还是不够这次要发的数量，就是真的缺货，报 shortfall 让调用方报错，不找别的货号
顶替——"变体"字段标的是"同一造型不同颜色"的分组，不是库存可以互换的依据（核对过真实数据，
不同颜色的货没法互相顶替发货），所以分摊只在同一个货号自己名下的订单里找，这里只负责
"给定一个型号能分摊出多少、还差多少"。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from openpyxl.worksheet.worksheet import Worksheet

from .column_utils import (
    HeaderNotFoundError,
    column_index_map,
    find_header_row,
    require_columns,
)

FIXED_HEADERS = ["订单号", "采购日期", "型号", "订单数量", "数量单位", "未出货数量"]
SUB_HEADER_LABEL = "出货时间"

EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def _to_date(value) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return (EXCEL_EPOCH + dt.timedelta(days=value)).date()
        except (OverflowError, OSError):
            return None
    if isinstance(value, str):
        for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None


@dataclass
class PurchaseRow:
    row_index: int
    order_no: str
    purchase_date: dt.date
    model: str
    order_qty: int
    initial_remaining: int
    consumed_this_run: int = 0

    @property
    def remaining(self) -> int:
        return self.initial_remaining - self.consumed_this_run


@dataclass
class Allocation:
    row: PurchaseRow
    quantity: int


@dataclass
class AllocationOutcome:
    allocations: list[Allocation] = field(default_factory=list)
    shortfall: int = 0


class DateColumnNotFoundError(Exception):
    pass


class PurchaseBook:
    def __init__(self, ws: Worksheet, progress_callback=None):
        self.ws = ws
        self.header_row = find_header_row(ws, FIXED_HEADERS, max_scan_rows=5, context="采购订单汇总表")
        self.sub_header_row = self.header_row + 1

        cols = column_index_map(ws, self.header_row)
        idx = require_columns(cols, FIXED_HEADERS, "采购订单汇总表")
        self.order_no_col = idx["订单号"]
        self.purchase_date_col = idx["采购日期"]
        self.model_col = idx["型号"]
        self.order_qty_col = idx["订单数量"]
        self.remaining_col = idx["未出货数量"]

        self.date_col_start = idx["数量单位"] + 1
        self.date_col_end = self.remaining_col - 1
        if self.date_col_end < self.date_col_start:
            raise HeaderNotFoundError("采购订单汇总表：在「数量单位」和「未出货数量」之间没找到日期列")

        self.rows: list[PurchaseRow] = []
        self.by_model: dict[str, list[PurchaseRow]] = {}
        self._load_rows(progress_callback)

    # ---- 读取 ----

    def _load_rows(self, progress_callback=None) -> None:
        self.rows.clear()
        self.by_model.clear()
        last_row = self.ws.max_row
        start_row = self.header_row + 2
        total = max(last_row - start_row + 1, 0)
        for done, r in enumerate(range(start_row, last_row + 1), start=1):
            if progress_callback is not None and (done % 200 == 0 or done == total):
                progress_callback(done, total)
            model = self.ws.cell(row=r, column=self.model_col).value
            if model is None or not str(model).strip():
                continue
            model = str(model).strip()

            order_no = self.ws.cell(row=r, column=self.order_no_col).value
            purchase_date = _to_date(self.ws.cell(row=r, column=self.purchase_date_col).value)
            order_qty = self.ws.cell(row=r, column=self.order_qty_col).value
            if purchase_date is None or not isinstance(order_qty, (int, float)):
                continue

            shipped_sum = 0
            for c in range(self.date_col_start, self.date_col_end + 1):
                v = self.ws.cell(row=r, column=c).value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    shipped_sum += v

            row_obj = PurchaseRow(
                row_index=r,
                order_no=str(order_no).strip() if order_no is not None else "",
                purchase_date=purchase_date,
                model=model,
                order_qty=int(order_qty),
                initial_remaining=int(order_qty) - int(shipped_sum),
            )
            self.rows.append(row_obj)
            self.by_model.setdefault(model, []).append(row_obj)

        for candidates in self.by_model.values():
            candidates.sort(key=lambda r: (r.purchase_date, r.row_index))

    # ---- 分摊 ----

    def allocate(self, model: str, quantity_needed: int) -> AllocationOutcome:
        outcome = AllocationOutcome()
        remaining_need = quantity_needed

        for row_obj in self.by_model.get(model, []):
            if remaining_need <= 0:
                break
            avail = row_obj.remaining
            if avail <= 0:
                continue
            take = min(avail, remaining_need)
            row_obj.consumed_this_run += take
            outcome.allocations.append(Allocation(row=row_obj, quantity=take))
            remaining_need -= take

        outcome.shortfall = remaining_need
        return outcome

    def find_date_column(self, target_date: dt.date) -> int | None:
        """只读查找，不创建——"发货计划汇总表更新"用来确认「采购订单分摊更新」是不是已经把
        这一天的分摊写进采购表了；找不到就说明这份采购表还不是"用过的"，不能拿来重放分摊
        结果（见 allocate_recorded 的说明）。
        """
        for c, d in self._real_date_columns():
            if d == target_date:
                return c
        return None

    def allocate_recorded(self, model: str, date_col: int, quantity_needed: int) -> AllocationOutcome:
        """不按"还剩多少余量"重新计算分摊，而是直接读 date_col 这一天的日期列里，这个货号
        名下每笔采购订单实际已经写了多少，按同样的先后顺序（采购日期从早到晚）依次分给
        这一条发货计划——"发货计划汇总表更新"专用：那份采购表传进来的时候已经是「采购订单
        分摊更新」写过的结果，这里只是把写死的分配结果读出来去匹配发货计划汇总表待定行该
        对应哪个（采购单号, 型号），不能再用 allocate() 那套"按余量现算"的逻辑——余量在
        写采购表那一步已经被这一批实际用掉了，这时候再按余量现算，同一货号有好几个不同
        ZD 批次时算出来的拆分顺序会跟当初实际写的对不上，重新算一遍余量够不够也没有意义
        （余量判断的时机是写采购表的时候，不是这里）。

        row_obj.consumed_this_run 在这个方法里的含义是"这一天记录的量已经被这次重放消耗
        了多少"，不是 allocate() 里"从余量里扣了多少"——这两套方法不会在同一个 PurchaseBook
        实例上混用，字段复用不会互相干扰。
        """
        outcome = AllocationOutcome()
        remaining_need = quantity_needed

        for row_obj in self.by_model.get(model, []):
            if remaining_need <= 0:
                break
            v = self.ws.cell(row=row_obj.row_index, column=date_col).value
            recorded = int(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0
            avail = recorded - row_obj.consumed_this_run
            if avail <= 0:
                continue
            take = min(avail, remaining_need)
            row_obj.consumed_this_run += take
            outcome.allocations.append(Allocation(row=row_obj, quantity=take))
            remaining_need -= take

        outcome.shortfall = remaining_need
        return outcome

    # ---- 写入 ----

    def _real_date_columns(self) -> list[tuple[int, dt.date | None]]:
        # SUM 区间（date_col_start..date_col_end）里混了几个不是真正"出货批次日期"的列
        # （比如表头是裸数字、行3副标题不是"出货时间"的几列），这些列依然要算在未出货数量的
        # 公式范围内（跟表格原有公式保持一致），但找"哪一列对应这个日期"、"新日期该插在哪"
        # 的时候，只能看行3副标题确实是"出货时间"的那些列，不然会被那几个裸数字表头误判成
        # 很晚/很早的日期。
        cols = []
        for c in range(self.date_col_start, self.date_col_end + 1):
            if self.ws.cell(row=self.sub_header_row, column=c).value == SUB_HEADER_LABEL:
                d = _to_date(self.ws.cell(row=self.header_row, column=c).value)
                cols.append((c, d))
        return cols

    def require_date_column(self, target_date: dt.date) -> int:
        """跟 find_date_column 一样只读查找，找不到就直接报错，不再像早期版本那样现场插入
        一新列——插入列即便小心处理了公式区间/格式/列宽，实际使用中还是出现过插入点右边的
        公式跟着错位的问题（见模块文档）。日期列必须已经存在于表格里，人工先把这一天的列
        加好，这里才会往里面写数据。
        """
        date_col = self.find_date_column(target_date)
        if date_col is None:
            raise DateColumnNotFoundError(
                f"采购订单汇总表里没有 {target_date.strftime('%Y-%m-%d')} 这一天的日期列——"
                "请先在表格里手动加好这一天的日期列，再重新更新"
            )
        return date_col

    def write_allocation(self, allocation: Allocation, date_col: int) -> None:
        cell = self.ws.cell(row=allocation.row.row_index, column=date_col)
        existing = cell.value
        new_value = (existing or 0) + allocation.quantity if isinstance(existing, (int, float)) else allocation.quantity
        cell.value = new_value
