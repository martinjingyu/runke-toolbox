"""发货计划汇总表：每个"采购单号+型号"在这张表里，发货时间="待定"的那些行就是这笔订单还在
工厂/没安排走的库存余量——**同一个采购单号+型号完全可能同时有好几行待定**，这些行加起来的
数量总和应该跟采购汇总表里这笔订单的未出货数量对得上，行与行之间没有另外的区别，都是同一批
"还没决定发去哪"的库存。所以分摊的时候不是"找唯一一行改"，而是"按行顺序（先到先扣）从这些
待定行里依次扣，扣完一行不够再扣下一行"，跟 purchase_book.py 里从多笔采购订单里依次分摊是
一个道理。

单独扣一行的时候，只有两种结果：
1. 这次要扣的数量 < 这一行剩余数量 -> 在这一行正上方插入一条新行，新行记这次实际扣掉的量；
   这一行自己的数量减掉这个量（箱数跟着重算），还是待定，可能在后面的分摊里继续被扣。
2. 这次要扣的数量 = 这一行剩余数量（正好扣完）-> 不插入新行，直接把这一行的发货时间从"待定"
   改成这次的日期，其它字段按新行的规则原地更新。

如果这个采购单号+型号名下所有待定行加起来都不够这次要发的数量，说明两张表本身的数据就对不上
（比如采购汇总表显示还有货，发货计划汇总表这边却没记全），直接报错，不猜、不硬写。

新行的内容策略是"整行照抄待定行，再覆盖需要改的几个字段"——不是逐个列名去挑要不要抄。这是
因为实测发现这张表里很多列（标签、CBM、总材重、总实重、重量、DP、FN sku 等）本身是公式，
而且是"引用自己这一行"的公式（比如"标签"是 =+B405，"总材重"是按本行箱数/长宽高算出来的），
新行如果只是把这些公式的值原样抄一份文字过去，要么把公式变成写死的旧数字（CBM/总材重这些
本该随新行箱数自动变化的数就再也不会变了），要么公式里的行号还指着旧行（比如插到新行里的
"=+B405"，新行明明是别的行号，会指错地方）。所以这里统一处理：整行复制过去，遇到公式就把
公式里"跟着这一行走"的单元格引用（形如字母+旧行号）换成新行号，再原样保留公式本身——这样
CBM/总材重这些依赖箱数的字段会在 Excel 里用新行的箱数自动重新算出正确结果，不用我们自己猜
换算公式。

**插入策略（重要）**：`apply_shipment()` 本身不碰 `self.ws`——每一笔分摊只在内存里记一下
"打算在哪一行上面插一条什么样的新行"、"哪一行的数量要改成多少"，行号也只在内存里虚拟地往下
推算（`_virtual_row`），不会真的调用 openpyxl 的 `insert_rows()`。真正把这些改动写回 worksheet
是 `materialize()` 的活，而且只对整张表做一次搬移。这是因为 openpyxl 的 `insert_rows()` 本身
是把插入点以下每一个单元格逐个搬到新位置，代价跟"插入点以下还有多少行"成正比；一批发货如果
有几十上百笔分摊要走"插入"这条路，每笔都真的插一次的话，代价是这些单笔代价的总和——实测在
一张 2.8 万行的真实表上，30 笔分摊能跑到近一分钟。改成先攒起来、最后对整张表只搬一遍，总代价
从"（插入次数）×（表还有多少行）"降到"表有多少行"这一个量级，不管一批发货有多少笔分摊，
插入成本基本固定。
"""
from __future__ import annotations

import bisect
import datetime as dt
import re
from copy import copy
from dataclasses import dataclass

from openpyxl.worksheet.worksheet import Worksheet

from .column_utils import column_index_map, find_header_row, require_columns

# 明确要清空、留给后续人工填写的字段
BLANK_FIELDS = ["仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"]

REQUIRED_HEADERS = ["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态", *BLANK_FIELDS]

PENDING_LABEL = "待定"
NEW_STATUS = "未发货"

_FORMULA_REF_RE = re.compile(r"([A-Za-z]{1,3})(\d+)")


class PendingRowNotFoundError(Exception):
    pass


class InconsistentQuantityError(Exception):
    pass


@dataclass
class ShipmentSummaryChange:
    kind: str  # "insert_above" 或 "convert_in_place"
    pending_row: int
    new_row: int | None  # kind=="insert_above" 时，新插入行的行号
    order_no: str
    model: str
    quantity: int
    box_capacity: float | None
    boxes: float | None
    boxes_exact: bool
    zd: str
    ship_date: dt.date
    pending_remaining_after: int | None  # kind=="insert_above" 时，待定行扣完之后剩多少


def _reindex_formula(formula: str, old_row: int, new_row: int) -> str:
    # 精确替换：只把公式里等于 old_row 的行号换成 new_row，其它行号原样保留。用在"把模板行的
    # 公式抄进一个全新插入的行"这种场景——新行是凭空冒出来的，公式里跟旧模板行本身无关的
    # 其它行引用不该被牵连着一起改。
    def repl(m: re.Match) -> str:
        letters, digits = m.group(1), m.group(2)
        if int(digits) == old_row:
            return f"{letters}{new_row}"
        return m.group(0)

    return _FORMULA_REF_RE.sub(repl, formula)


class ShipmentSummaryBook:
    def __init__(self, ws: Worksheet):
        self.ws = ws
        self.header_row = find_header_row(ws, REQUIRED_HEADERS, max_scan_rows=10)
        cols = column_index_map(ws, self.header_row)
        self.col = require_columns(cols, REQUIRED_HEADERS, "发货计划汇总表")
        # openpyxl 的 ws.max_row / ws.max_column 不是缓存属性——每次访问都要把底层存储重新扫一遍
        # 找最大行号/列号（在几万行的表上，一次访问就是几万次内部循环）。这张表在这个 book 实例
        # 的生命周期里只会往下插行，不会插列，所以这里缓存一份、自己维护。
        self._max_row = ws.max_row
        self._max_column = ws.max_column
        # (采购单号, 型号) -> 待定行号列表（行号升序，对应先到先扣的顺序）。之前每次调用都要把
        # 整张表从头扫到尾，一批发货几十上百次分摊等于反复全表扫描，是这个工具最大的耗时来源
        # 之一。这里在加载时只扫一遍表建好索引，之后转正/插入只增量维护，不用再重新扫表。
        self._pending_index: dict[tuple[str, str], list[int]] = {}
        # 下面这三个是"这一批分摊攒下来、还没真正写进 worksheet"的改动，见 materialize()：
        #   _insert_offsets：按原始行号排序的"在这一行上面插了一条新行"事件列表（可以有重复——
        #     同一行在这一批里被拆分好几次的话，就会出现好几条）。任意原始行号 r 经过这一批
        #     全部插入之后的"虚拟当前行号"，用 r + bisect_right(_insert_offsets, r) 算，不用
        #     真的挪动 worksheet 就能拿到准确答案（原理跟 diff.py 里 row_of/_shift_from 一样，
        #     只是这里把它内置进 ShipmentSummaryBook 自己，覆盖的是"这一批摸过的所有行"而不是
        #     只有 diff 预览关心的那一小撮）。
        #   _field_overrides：原始行号 -> {列号: 新值}，这一行最终落地时要覆盖成什么（转正的
        #     全部字段，或者拆分之后剩余数量/箱数）。
        #   _queued_inserts：原始行号（插入点所在的那一行）-> 要插在它上面的新行内容列表
        #     （字段覆盖 dict，跟 _field_overrides 是同一种形状），按插入发生的先后顺序排列。
        self._insert_offsets: list[int] = []
        self._field_overrides: dict[int, dict[int, object]] = {}
        self._queued_inserts: dict[int, list[dict[int, object]]] = {}
        self._build_pending_index()

    def _build_pending_index(self) -> None:
        self._pending_index = {}
        c_order = self.col["采购单号"]
        c_model = self.col["型号"]
        c_ship = self.col["发货时间"]
        c_status = self.col["状态"]
        for r in range(self.header_row + 1, self._max_row + 1):
            if (
                self.ws.cell(row=r, column=c_ship).value == PENDING_LABEL
                # 待定库存必须还是"未发货"状态才真的能被分摊——真实表里踩过坑：有些行已经被
                # 标成"已取消"/"无库存"，但当时改状态的人忘了把"发货时间"这一列也从"待定"
                # 改回去，只看"发货时间"会把这些早就作废的行错当成还能用的库存去扣，把货错发
                # 到不该发的地方。两个条件都满足才算数。
                and self.ws.cell(row=r, column=c_status).value == NEW_STATUS
            ):
                order_no = self.ws.cell(row=r, column=c_order).value
                model = self.ws.cell(row=r, column=c_model).value
                self._pending_index.setdefault((order_no, model), []).append(r)

    def _remove_from_pending_index(self, order_no: str, model: str, row: int) -> None:
        rows = self._pending_index.get((order_no, model))
        if rows is not None and row in rows:
            rows.remove(row)

    def _virtual_row(self, original_row: int) -> int:
        """这一批分摊到目前为止，原始行号 original_row 现在"虚拟地"在第几行——不碰
        worksheet，纯粹靠"它上面已经插了几条新行"算出来，见 __init__ 里 _insert_offsets
        的说明。
        """
        return original_row + bisect.bisect_right(self._insert_offsets, original_row)

    def pending_rows(self, order_no: str, model: str) -> list[int]:
        """按行顺序返回这个采购单号+型号名下所有"发货时间=待定 且 状态=未发货"的行号
        （可能不止一个）。直接查 __init__ 时建好的索引，不用每次都扫表。返回的是"原始行号"
        （这一批分摊还没 materialize() 之前，行号本身不会变，见类文档）。
        """
        return list(self._pending_index.get((order_no, model), []))

    def find_pending_row(self, order_no: str, model: str) -> int | None:
        """随便找一行待定行（不保证是哪一行，也不保证是唯一一行）——只用来做"这个采购单号+
        型号到底存不存在待定库存"这种存在性判断，分摊逻辑不应该用这个，见 apply_shipment。
        """
        rows = self.pending_rows(order_no, model)
        return rows[0] if rows else None

    def total_pending_quantity(self, order_no: str, model: str) -> int:
        return sum(self.read_quantity(r) for r in self.pending_rows(order_no, model))

    def read_quantity(self, row: int) -> int:
        # 这一批里已经改过这一行数量的话（比如前一笔分摊刚把它拆小了），要看内存里记的最新值，
        # 不能去 worksheet 上读——materialize() 之前 worksheet 上那一格还是没动过的旧值。
        override = self._field_overrides.get(row, {}).get(self.col["数量"])
        if override is not None:
            return int(override)

        value = self.ws.cell(row=row, column=self.col["数量"]).value
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return int(value)
        if isinstance(value, str) and value.startswith("="):
            # "数量"本身是公式（一般是 箱数*箱容），公式结果读不到，改成自己拿箱数*箱容算
            boxes = self.ws.cell(row=row, column=self.col["箱数"]).value
            capacity = self.ws.cell(row=row, column=self.col["箱容"]).value
            if isinstance(boxes, (int, float)) and isinstance(capacity, (int, float)):
                return int(boxes * capacity)
        raise ValueError(f"发货计划汇总表第 {row} 行的「数量」既不是数字也不是能识别的公式，读不出来")

    def apply_shipment(
        self, order_no: str, model: str, quantity: int, zd: str, ship_date: dt.date
    ) -> list[ShipmentSummaryChange]:
        """把 quantity 这么多货，从这个采购单号+型号名下的待定行里依次扣掉，可能一次扣完
        一行也可能要扣好几行（先到先扣），返回按处理顺序排列的改动列表。这一步只在内存里记账，
        不碰 worksheet，见类文档"插入策略"——真正写回表格要调用 materialize()。
        """
        total_available = self.total_pending_quantity(order_no, model)
        if not self.pending_rows(order_no, model):
            raise PendingRowNotFoundError(
                f"发货计划汇总表里找不到采购单号「{order_no}」+ 型号「{model}」发货时间=待定 的行"
            )
        if quantity > total_available:
            # 正常走到这里之前，采购汇总表那边已经确认这个采购单号+型号有足够的未出货数量，
            # 理论上不该出现要发的比这边所有待定行加起来还多——真出现了，说明两张表本身就对
            # 不上（比如发货计划汇总表这边库存没记全），不该硬着头皮把数字覆盖过去，宁可停
            # 下来报错让人看。
            raise InconsistentQuantityError(
                f"发货计划汇总表里，采购单号「{order_no}」+ 型号「{model}」名下所有待定行的数量"
                f"加起来只有 {total_available}，但要写入的数量是 {quantity}，比这还多，数据对不上，"
                f"不能写入"
            )

        changes: list[ShipmentSummaryChange] = []
        remaining_need = quantity
        while remaining_need > 0:
            # 每一轮都重新找"当前还有余量的第一行"，不能缓存行号列表——同一行可能在这一批里
            # 被前一轮拆小过，read_quantity 会看到最新的（内存里的）余量。
            row = next(
                (r for r in self.pending_rows(order_no, model) if self.read_quantity(r) > 0),
                None,
            )
            if row is None:
                # total_available 已经在前面校验过够，理论上不会走到这里；真出现了说明前面的
                # 校验和这里的实际扣减对不上，同样不该瞎猜，直接报错。
                raise InconsistentQuantityError(
                    f"发货计划汇总表里，采购单号「{order_no}」+ 型号「{model}」的待定行在处理过程中"
                    f"意外用完了，还差 {remaining_need} 没能分摊，数据可能有问题，需要人工核对"
                )
            row_qty = self.read_quantity(row)
            take = min(row_qty, remaining_need)
            changes.append(self._consume_row(row, take, order_no, model, zd, ship_date))
            remaining_need -= take

        return changes

    def _explicit_and_blank_fields(
        self, order_no: str, model: str, quantity: int, boxes, zd: str, ship_date: dt.date, status: str
    ) -> dict[int, object]:
        fields: dict[int, object] = {
            self.col["采购单号"]: order_no,
            self.col["型号"]: model,
            self.col["数量"]: quantity,
            self.col["箱数"]: boxes,
            self.col["ZD"]: zd,
            self.col["发货时间"]: dt.datetime.combine(ship_date, dt.time()),
            self.col["状态"]: status,
        }
        for field in BLANK_FIELDS:
            col = self.col.get(field)
            if col is not None:
                fields[col] = None
        return fields

    def _consume_row(
        self, pending_row: int, quantity: int, order_no: str, model: str, zd: str, ship_date: dt.date
    ) -> ShipmentSummaryChange:
        """从这一行待定库存里扣 quantity（调用方保证 0 < quantity <= 这一行当前数量）。
        只更新内存里的记账，不碰 worksheet。
        """
        pending_qty = self.read_quantity(pending_row)
        box_capacity = self.ws.cell(row=pending_row, column=self.col["箱容"]).value
        boxes, boxes_exact = _compute_boxes(quantity, box_capacity)

        if quantity < pending_qty:
            # new_row 是"这一行现在虚拟地在第几行"——插入发生在它上面，新行就落在这个位置，
            # 原来这一行的内容（含之后新算出的剩余数量）整体往下挪一位，变成 new_row + 1。
            new_row = self._virtual_row(pending_row)
            template_row = new_row + 1

            insert_fields = self._explicit_and_blank_fields(
                order_no, model, quantity, boxes, zd, ship_date, NEW_STATUS
            )
            self._queued_inserts.setdefault(pending_row, []).append(insert_fields)
            bisect.insort(self._insert_offsets, pending_row)

            remaining = pending_qty - quantity
            remaining_boxes, _ = _compute_boxes(remaining, box_capacity)
            overrides = self._field_overrides.setdefault(pending_row, {})
            overrides[self.col["数量"]] = remaining
            overrides[self.col["箱数"]] = remaining_boxes

            return ShipmentSummaryChange(
                kind="insert_above",
                pending_row=template_row,
                new_row=new_row,
                order_no=order_no,
                model=model,
                quantity=quantity,
                box_capacity=box_capacity,
                boxes=boxes,
                boxes_exact=boxes_exact,
                zd=zd,
                ship_date=ship_date,
                pending_remaining_after=remaining,
            )

        # quantity == pending_qty：正好扣完这一行，原地转正（不插入新行，这一行的虚拟行号
        # 不会因为这次分摊本身而变化，只会被"别的行"产生的插入事件影响，同样靠 _virtual_row 算）。
        final_row = self._virtual_row(pending_row)
        overrides = self._field_overrides.setdefault(pending_row, {})
        overrides.update(
            self._explicit_and_blank_fields(order_no, model, quantity, boxes, zd, ship_date, NEW_STATUS)
        )
        # 这一行发货时间从"待定"变成了具体日期，不再是待定库存，从索引里摘掉，不然后面
        # 同一个采购单号+型号再来一笔分摊，会把这一行当成还能扣的库存重复用。
        self._remove_from_pending_index(order_no, model, pending_row)
        return ShipmentSummaryChange(
            kind="convert_in_place",
            pending_row=final_row,
            new_row=None,
            order_no=order_no,
            model=model,
            quantity=quantity,
            box_capacity=box_capacity,
            boxes=boxes,
            boxes_exact=boxes_exact,
            zd=zd,
            ship_date=ship_date,
            pending_remaining_after=None,
        )

    # ---- 把攒下来的改动真正写回 worksheet ----

    def materialize(self) -> None:
        """把这一批 apply_shipment() 攒下来的所有插入/字段修改，一次性写回 worksheet。
        必须在读取/保存 self.ws 的最终内容之前调用一次（比如 diff.py 要生成"变化后"快照、
        或者调用方要 wb.save() 之前）；一批发货计划只需要调用一次，不管这一批有多少笔分摊。

        算法：从当前表格最后一行开始往前处理到表头，对每一行 r 算出它这一批下来最终落在
        第几行（dest = r 加上"它上面这一批一共插了几条新行"），只有 dest != r（真的要挪位置）
        或者这一行本身有改动/有新行要插在它上面时，才读它、写它——大部分行完全没被这一批碰到，
        直接跳过，一分钱不花。从大到小处理是为了保证："读某一行原始内容"这一步一定发生在
        "别的行的内容被写进这一行"之前，不会读到已经被覆盖掉的脏数据（跟 purchase_book.py 里
        _shift_column_dimensions 那种"从后往前处理"的写法是同一个道理）。
        """
        if not self._insert_offsets and not self._field_overrides:
            return  # 这一批什么都没改，不用碰表格

        total_inserts = len(self._insert_offsets)
        if total_inserts == 0:
            # 只有原地转正、没有任何插入——不用挪任何一行，直接把字段覆盖写到原地
            for row, overrides in self._field_overrides.items():
                for col, value in overrides.items():
                    self.ws.cell(row=row, column=col).value = value
            self._finish_materialize()
            return

        old_max_row = self._max_row
        for r in range(old_max_row, self.header_row, -1):
            shift = bisect.bisect_right(self._insert_offsets, r)
            dest = r + shift
            queued = self._queued_inserts.get(r)
            overrides = self._field_overrides.get(r)

            if dest == r and queued is None and overrides is None:
                continue  # 这一行完全没被这一批影响，跳过

            values, styles, row_dim = self._snapshot_row(r)

            if queued:
                k = len(queued)
                for i, insert_fields in enumerate(queued):
                    self._write_row(dest - k + i, r, values, styles, row_dim, insert_fields, is_new_row=True)

            if dest != r or overrides:
                self._write_row(dest, r, values, styles, row_dim, overrides or {}, is_new_row=False)

        self._max_row = old_max_row + total_inserts
        self._finish_materialize()

    def _finish_materialize(self) -> None:
        self._insert_offsets = []
        self._field_overrides = {}
        self._queued_inserts = {}
        # 这一批插入/转正之后，待定行的行号和"还有哪些行待定"都变了，重新扫一遍建索引——
        # 这一步本身也是 O(表格行数)，跟 materialize() 主体是同一个量级，不算额外的负担。
        self._build_pending_index()

    def _snapshot_row(self, row: int):
        values: dict[int, object] = {}
        styles: dict[int, object] = {}
        for c in range(1, self._max_column + 1):
            cell = self.ws.cell(row=row, column=c)
            values[c] = cell.value
            if cell.has_style:
                # 不要通过 cell.font/cell.fill/... 这些高层属性去读/抄样式——每个单元格实际
                # 存的只是 cell._style 这个 9 个整数的 StyleArray（字体/填充/边框……分别是
                # 指向共享样式表的一个下标），cell.font 这类属性每次访问都会现查一遍共享表、
                # 包一层 StyleProxy 返回；对着 StyleProxy 再 copy() 更是很贵的操作（内部会走
                # to_tree/from_tree 序列化一圈）。直接复制 _style 这个小整数数组，是 openpyxl
                # 自己在 WorksheetCopy 里搬一整个工作表时用的同一种手法（见
                # openpyxl/worksheet/copier.py），既快又是官方认可的搬运方式；用 copy() 是因为
                # StyleArray 这几个下标不能多个单元格共享同一个数组实例——后面谁改了自己的
                # 字体，会连带把共享这个数组的其它单元格也改了。
                styles[c] = copy(cell._style)
        dim = self.ws.row_dimensions.get(row)
        row_dim = (dim.height, dim.hidden, dim.outlineLevel, dim.collapsed) if dim is not None else None
        return values, styles, row_dim

    def _write_row(
        self,
        dest_row: int,
        src_row: int,
        src_values: dict[int, object],
        src_styles: dict[int, tuple],
        src_row_dim,
        field_overrides: dict[int, object],
        is_new_row: bool,
    ) -> None:
        for c in range(1, self._max_column + 1):
            value = src_values.get(c)
            if isinstance(value, str) and value.startswith("="):
                if is_new_row:
                    # 全新插入的行：只把公式里"指向模板行自己"的引用换成这一新行的行号，
                    # 公式里跟模板行无关的其它行引用原样保留——这一行是凭空冒出来的，不该
                    # 牵连其它引用。
                    value = _reindex_formula(value, src_row, dest_row)
                else:
                    # 只是把一行原有内容整体挪到新位置：公式里任何行引用（不管是不是等于
                    # src_row 自己）只要落在"这一批插入影响的范围内"，都要跟着改成挪位置后的
                    # 结果——既覆盖"引用自己这一行"的公式，也覆盖表底 SUBTOTAL 那种跨很多行
                    # 的区间引用。
                    value = self._reindex_formula_bulk(value)
            dest_cell = self.ws.cell(row=dest_row, column=c)
            dest_cell.value = value
            style = src_styles.get(c)
            if style is not None:
                # 直接整个 StyleArray 赋值——见 _snapshot_row 里的说明，这一步等于把字体/
                # 填充/边框/对齐/数字格式/保护属性一次性全部搬过去，不用逐个属性再各查一遍。
                dest_cell._style = copy(style)

        for col, value in field_overrides.items():
            self.ws.cell(row=dest_row, column=col).value = value

        if src_row_dim is not None:
            height, hidden, outline, collapsed = src_row_dim
            dim = self.ws.row_dimensions[dest_row]
            dim.height = height
            dim.hidden = hidden
            dim.outlineLevel = outline
            dim.collapsed = collapsed

    def _reindex_formula_bulk(self, formula: str) -> str:
        def repl(m: re.Match) -> str:
            letters, digits = m.group(1), m.group(2)
            old_row = int(digits)
            return f"{letters}{self._virtual_row(old_row)}"

        return _FORMULA_REF_RE.sub(repl, formula)


def _compute_boxes(quantity: int, box_capacity) -> tuple[float | None, bool]:
    if not isinstance(box_capacity, (int, float)) or box_capacity <= 0:
        return None, False
    boxes = quantity / box_capacity
    exact = float(boxes).is_integer()
    return (int(boxes) if exact else boxes), exact
