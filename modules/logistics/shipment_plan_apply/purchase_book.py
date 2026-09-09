"""采购订单汇总表：每一行是"一笔采购订单里的一个型号"，后面跟着一长串"日期列"（哪一批次
出货，就在那一列写发出的数量），"未出货数量"是表格里已经有的本地公式：
    =+<订单数量列><行>-SUM(<第一个日期列><行>:<最后一个日期列><行>)
本工具只需要把这次发货的数量填进正确的日期列——公式本身在 Excel 打开时会自己算出新的
"未出货数量"，不用我们手动改那一格。

但如果要写入的日期还没有对应的列，需要真的在表格中间插入一新列（按日期顺序插在合适的位置，
而不是随手加在最后，方便人工阅读）。openpyxl 的 insert_cols 只会把已有单元格的内容搬到新位置，
并不会像 Excel 那样自动把公式里的区间引用（比如 SUM(K4:BU4)）跟着调整——如果不管这件事，
插入一列之后所有行的"未出货数量"公式实际引用的区间就会跟数据错位，算出错的数字还看不出来。
所以每次插入日期列之后，都要把所有行的"未出货数量"公式重新按当前的实际列范围写一遍。

分摊规则：同一个货号（型号）可能对应好几笔采购订单，按采购日期从早到晚，先把早的写满，
写满了（未出货数量到 0）就换下一笔。如果找不到货号对应的订单，或者所有能找到的订单加起来
的未出货数量还是不够这次要发的数量，就是真的缺货，报 shortfall 让调用方报错，不找别的货号
顶替——"变体"字段标的是"同一造型不同颜色"的分组，不是库存可以互换的依据（核对过真实数据，
不同颜色的货没法互相顶替发货），所以分摊只在同一个货号自己名下的订单里找，这里只负责
"给定一个型号能分摊出多少、还差多少"。
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from copy import copy

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

from .column_utils import (
    HeaderNotFoundError,
    column_index_map,
    find_header_row,
    require_columns,
    unmerge_overlapping_columns,
)

_FORMULA_REF_RE = re.compile(r"([A-Za-z]{1,3})(\d+)")


def _reindex_formula_column(formula: str, old_col: int, new_col: int) -> str:
    # 精确替换：只把公式里等于 old_col 这一列的引用换成 new_col，其它列号原样保留——用在
    # "把一个格子原样搬到隔壁列"这种场景，不是"整个区间跟着挪"（那个是 _rewrite_remaining_
    # formulas 在插入完之后统一按新的列范围重写，不靠这里的逐格搬运）。
    def repl(m: re.Match) -> str:
        letters, digits = m.group(1), m.group(2)
        if column_index_from_string(letters.upper()) == old_col:
            return f"{get_column_letter(new_col)}{digits}"
        return m.group(0)

    return _FORMULA_REF_RE.sub(repl, formula)

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

    def find_or_create_date_column(self, target_date: dt.date) -> int:
        real_cols = self._real_date_columns()

        for c, d in real_cols:
            if d == target_date:
                return c

        insert_at = (real_cols[-1][0] + 1) if real_cols else (self.date_col_end + 1)
        for c, d in real_cols:
            if d is not None and d > target_date:
                insert_at = c
                break

        # 插入之前先记下要照抄格式的那一列（本来在插入点左边的那一列，插入之后位置不变，
        # 还是原来的列号；插入点正好是日期区间最左边、左边没有可抄的日期列时，就抄插入之后
        # 落在右边的那一列）——旧列右移之后这个"左边列号"不会跟着变，插入完了正好能直接用。
        style_source = insert_at - 1 if insert_at - 1 >= self.date_col_start else None

        # 不用 openpyxl 自带的 insert_cols()——它不管插入点在哪，内部都要把整张表 _cells
        # 字典里的全部单元格先按列排一次序（跟 shipment_summary.py 里绕开 insert_rows() 是
        # 同一个坑）；这张表 ws.max_column 常年被撑到 Excel 列数上限（某个很远的格子只是
        # 留了点格式的假象，不是真的有这么多列有数据），实测这一步能跑到 166 秒。这里改成
        # 只在真正有数据的列范围内（insert_at 到 remaining_col）搬。
        #
        # 真实表格里这一片区域出现过历史遗留的合并单元格（比如带过批注的格子）——合并区域里
        # 非左上角的格子是 openpyxl 的 MergedCell，连 .value 都赋不了值，搬运之前得先拆开，
        # 见 unmerge_overlapping_columns 的说明。
        unmerge_overlapping_columns(self.ws, insert_at, self.remaining_col + 1)
        self._shift_columns_right(insert_at, self.remaining_col)
        self._shift_column_dimensions(insert_at)
        style_col = style_source if style_source is not None else insert_at + 1
        self._copy_column_style(insert_at, style_col)
        self._copy_row1_total_formula(insert_at, style_col)

        self.ws.cell(row=self.header_row, column=insert_at, value=dt.datetime.combine(target_date, dt.time()))
        self.ws.cell(row=self.sub_header_row, column=insert_at, value=SUB_HEADER_LABEL)

        self.date_col_end += 1
        if insert_at <= self.remaining_col:
            self.remaining_col += 1
        self._rewrite_remaining_formulas()
        self._sync_auto_filter()

        return insert_at

    def _shift_columns_right(self, insert_at: int, real_last_col: int) -> None:
        """把 [insert_at, real_last_col] 这个真正有数据的列区间整体右移一列，给新插入的
        日期列腾位置，insert_at 这一列本身腾空。只在这个有限范围内搬（不是整张表 max_column），
        从最右边的列开始往左处理，不然后面的赋值会覆盖掉还没读出来的旧值（跟
        shipment_summary.py 里 _push_tail_row_down 从后往前处理是同一个道理，只是行列换了个
        方向）。这张表的日期列/未出货数量列实测都是普通数值/在插入完之后会被
        _rewrite_remaining_formulas() 整个重写的公式，不需要处理"公式里引用别的列"这种情况，
        但为了不留坑，遇到公式还是按"引用自己这一列"的规则挪一下，不假设它一定是纯数值。
        """
        last_row = self.ws.max_row
        for c in range(real_last_col, insert_at - 1, -1):
            dest_c = c + 1
            for r in range(1, last_row + 1):
                src_cell = self.ws.cell(row=r, column=c)
                value = src_cell.value
                if isinstance(value, str) and value.startswith("="):
                    value = _reindex_formula_column(value, c, dest_c)
                dest_cell = self.ws.cell(row=r, column=dest_c)
                dest_cell.value = value
                if src_cell.has_style:
                    dest_cell._style = copy(src_cell._style)
        # 腾空的这一列（insert_at）还留着搬走之前的旧值，得显式清掉，不然后面写新日期列的时候
        # 底下的数据行会看起来像是这一列本来就有历史数据。
        for r in range(1, last_row + 1):
            self.ws.cell(row=r, column=insert_at).value = None

    def _sync_auto_filter(self) -> None:
        # 插入新的日期列会让表格整体变宽一列，但 Excel 的筛选范围是写死在文件里的固定区间，
        # 不会自动跟着扩——跟 shipment_summary.py 里 sync_auto_filter 是同一个坑，那边是
        # 行范围过期，这边是列范围过期。只扩大、不缩小。
        current_ref = self.ws.auto_filter.ref
        target_min_row, target_min_col = self.header_row, 1
        target_max_row, target_max_col = self.ws.max_row, self.remaining_col
        if current_ref:
            min_col, min_row, max_col, max_row = range_boundaries(current_ref)
            target_min_row = min(target_min_row, min_row)
            target_min_col = min(target_min_col, min_col)
            target_max_row = max(target_max_row, max_row)
            target_max_col = max(target_max_col, max_col)
        self.ws.auto_filter.ref = (
            f"{get_column_letter(target_min_col)}{target_min_row}:"
            f"{get_column_letter(target_max_col)}{target_max_row}"
        )

    def _shift_column_dimensions(self, inserted_at: int) -> None:
        # 从最右边的列开始往左处理，不然后面的赋值会覆盖掉还没读出来的旧值（跟
        # shipment_summary.py 里 _push_tail_row_down 从后往前处理是同一个道理，只是行列换了个方向）。
        existing_cols = sorted(
            (
                idx
                for letter in self.ws.column_dimensions
                if (idx := column_index_from_string(letter)) >= inserted_at
            ),
            reverse=True,
        )
        for idx in existing_cols:
            src = self.ws.column_dimensions[get_column_letter(idx)]
            dest = self.ws.column_dimensions[get_column_letter(idx + 1)]
            dest.width = src.width
            dest.hidden = src.hidden
            dest.outlineLevel = src.outlineLevel
            dest.collapsed = src.collapsed
        inserted_letter = get_column_letter(inserted_at)
        if inserted_letter in self.ws.column_dimensions:
            del self.ws.column_dimensions[inserted_letter]

    def _copy_column_style(self, dest_col: int, src_col: int) -> None:
        if self.ws.max_row is None:
            return
        src_letter = get_column_letter(src_col)
        if src_letter in self.ws.column_dimensions:
            src_dim = self.ws.column_dimensions[src_letter]
            dest_dim = self.ws.column_dimensions[get_column_letter(dest_col)]
            dest_dim.width = src_dim.width
        for r in range(1, self.ws.max_row + 1):
            src_cell = self.ws.cell(row=r, column=src_col)
            dest_cell = self.ws.cell(row=r, column=dest_col)
            # 不管 src_cell 有没有"显式样式"（has_style）都要复制，不能跳过——插入点这一列
            # 在插入之前可能残留着这张表 ws.max_column 被撑得远超实际数据范围导致的杂散格式
            # （见上面 unmerge_overlapping_columns 附近的说明），如果 src 没样式就 continue
            # 跳过不写，dest 会保留这些杂散格式不被清掉，插进来的新日期列格式会跟旁边列对不上
            # （实测：日期头那一格会显示成裸数字，不是 m月d日 这种正常格式）。直接复制 _style
            # 这个小整数数组（不经过 font/fill/... 这些高层属性），是 openpyxl 官方认可的搬运
            # 方式，也比逐个属性复制快得多——跟 shipment_summary.py 里 _copy_row_with_reindex
            # 用的是同一个手法。
            dest_cell._style = copy(src_cell._style)

    def _copy_row1_total_formula(self, dest_col: int, src_col: int) -> None:
        # 第 1 行（表头上面那一行）是每一列自己的求和行（比如 =SUBTOTAL(9,BT4:BT1000732)），
        # 新插入的日期列本该也有一份、只是列号换成自己——_copy_column_style 只搬样式不搬公式，
        # 不补这一步的话新插入的这一列在第 1 行会一直是空的，Excel 里这一列的求和格看着像是
        # 漏填了。src_col 的公式如果本身就不是"引用自己这一列"的求和公式（比如凑巧是空的、
        # 或者是别的写法），就不猜、不硬造一个，保持这一格空着，跟 src 没有可抄的东西时一样。
        src_formula = self.ws.cell(row=1, column=src_col).value
        if isinstance(src_formula, str) and src_formula.startswith("="):
            self.ws.cell(row=1, column=dest_col).value = _reindex_formula_column(src_formula, src_col, dest_col)

    def _rewrite_remaining_formulas(self) -> None:
        qty_letter = get_column_letter(self.order_qty_col)
        start_letter = get_column_letter(self.date_col_start)
        end_letter = get_column_letter(self.date_col_end)
        for row_obj in self.rows:
            r = row_obj.row_index
            formula = f"=+{qty_letter}{r}-SUM({start_letter}{r}:{end_letter}{r})"
            self.ws.cell(row=r, column=self.remaining_col, value=formula)

    def write_allocation(self, allocation: Allocation, date_col: int) -> None:
        cell = self.ws.cell(row=allocation.row.row_index, column=date_col)
        existing = cell.value
        new_value = (existing or 0) + allocation.quantity if isinstance(existing, (int, float)) else allocation.quantity
        cell.value = new_value
