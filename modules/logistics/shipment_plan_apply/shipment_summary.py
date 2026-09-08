"""发货计划汇总表：每个"采购单号+型号"在这张表里，发货时间="待定"的那些行就是这笔订单还在
工厂/没安排走的库存余量——**同一个采购单号+型号完全可能同时有好几行待定**，这些行加起来的
数量总和应该跟采购汇总表里这笔订单的未出货数量对得上，行与行之间没有另外的区别，都是同一批
"还没决定发去哪"的库存。所以分摊的时候不是"找唯一一行改"，而是"按行顺序（先到先扣）从这些
待定行里依次扣，扣完一行不够再扣下一行"，跟 purchase_book.py 里从多笔采购订单里依次分摊是
一个道理。

单独扣一行的时候，只有两种结果：
1. 这次要扣的数量 < 这一行剩余数量 -> 待定行原地不动，只把数量/箱数减掉这次扣的量（还是
   待定，可能在后面的分摊里继续被扣），另外单独插一行"已发货"记录这次实际扣掉的量。
2. 这次要扣的数量 = 这一行剩余数量（正好扣完）-> 不插入新行，直接把这一行的发货时间从"待定"
   改成这次的日期，其它字段按新行的规则原地更新。

如果这个采购单号+型号名下所有待定行加起来都不够这次要发的数量，说明两张表本身的数据就对不上
（比如采购汇总表显示还有货，发货计划汇总表这边却没记全），直接报错，不猜、不硬写。

**新插入的"已发货"行放在哪里（重要，性能关键）**：不是插在被扣的那条待定行正上方，而是永远
插在整张表最下面——如果表格最后一行是"合计"行（比如「数量」列是
`=SUBTOTAL(9,E6:E27947)` 这种区间汇总公式，"采购单号"/"型号"都是空的），就插在合计行上面
（顺便把合计公式的区间边界往下扩一位，让新记录也算进合计里）；如果表格最后一行就是普通数据行，
直接接在表格末尾，不用插入。已经跟用户确认过这个取舍：新记录不需要跟它对应的待定行挨在一起，
只要保证"这个 SKU 名下最后一条待定记录"始终是对的就行。

这么做是因为 openpyxl 的 `insert_rows()` 本身很贵，而且贵的地方不止是"插入点以下的单元格要
逐个挪位置"——它内部搬单元格之前会把整张表 `_cells` 字典的全部坐标先排一次序（源码见
`openpyxl/worksheet/worksheet.py` 的 `_move_cells`），不管插入点在哪，这一步都要碰一遍全表，
所以哪怕插入点已经很靠后，直接调用它也不便宜。这里改成插在待定行原来的位置上完全不用挪任何
东西（真的是原地改两个数），插在表格最下面的新行最多只需要搬"合计行"这一行（如果有的话）——
这两种情况都是常数级代价，跟这一批发货有多少笔分摊、待定行分散在表格哪里都没关系，也不用像
之前那样先攒一批改动最后统一"物化"到表格里（见 git 历史），`apply_shipment()` 现在就是
直接写、立刻生效。

新行的内容策略是"整行照抄待定行，再覆盖需要改的几个字段"——不是逐个列名去挑要不要抄。这是
因为实测发现这张表里很多列（标签、CBM、总材重、总实重、重量、DP、FN sku 等）本身是公式，
而且是"引用自己这一行"的公式（比如"标签"是 =+B405，"总材重"是按本行箱数/长宽高算出来的），
新行如果只是把这些公式的值原样抄一份文字过去，要么把公式变成写死的旧数字（CBM/总材重这些
本该随新行箱数自动变化的数就再也不会变了），要么公式里的行号还指着旧行（比如插到新行里的
"=+B405"，新行明明是别的行号，会指错地方）。所以这里统一处理：整行复制过去，遇到公式就把
公式里"跟着这一行走"的单元格引用（形如字母+旧行号）换成新行号，再原样保留公式本身——这样
CBM/总材重这些依赖箱数的字段会在 Excel 里用新行的箱数自动重新算出正确结果，不用我们自己猜
换算公式。
"""
from __future__ import annotations

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
    new_row: int | None  # kind=="insert_above" 时，新插入行的行号（在表格最下面，不挨着 pending_row）
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


def _grow_tail_formula(formula: str, old_tail_row: int, new_tail_row: int) -> str:
    # 合计行的公式一般是"引用一个数据区间"（比如 SUBTOTAL(9,E6:E27947)），区间的结束边界
    # 通常正好是"合计行自己往上一行"（也就是最后一条真实数据）；合计行往下挪了之后，这个
    # 边界要跟着往后扩，把新插进来、原来合计行所在位置的那些新记录也算进去。同时也顺手处理
    # 合计行自己如果有"引用自己这一行"的公式（不常见，但跟别的地方保持一个逻辑）。
    def repl(m: re.Match) -> str:
        letters, digits = m.group(1), m.group(2)
        row_num = int(digits)
        if row_num >= old_tail_row or row_num == old_tail_row - 1:
            return f"{letters}{row_num + (new_tail_row - old_tail_row)}"
        return m.group(0)

    return _FORMULA_REF_RE.sub(repl, formula)


class ShipmentSummaryBook:
    def __init__(self, ws: Worksheet):
        self.ws = ws
        self.header_row = find_header_row(ws, REQUIRED_HEADERS, max_scan_rows=10, context="发货计划汇总表")
        cols = column_index_map(ws, self.header_row)
        self.col = require_columns(cols, REQUIRED_HEADERS, "发货计划汇总表")
        # openpyxl 的 ws.max_row / ws.max_column 不是缓存属性——每次访问都要把底层存储重新扫一遍
        # 找最大行号/列号（在几万行的表上，一次访问就是几万次内部循环）。这张表在这个 book 实例
        # 的生命周期里只会往下插行，不会插列，所以这里缓存一份、自己维护。
        self._max_row = ws.max_row
        self._max_column = ws.max_column
        # (采购单号, 型号) -> 待定行号列表（行号升序，对应先到先扣的顺序）。之前每次调用都要把
        # 整张表从头扫到尾，一批发货几十上百次分摊等于反复全表扫描，是这个工具最大的耗时来源
        # 之一。这里在加载时只扫一遍表建好索引；待定行本身以后再也不会挪位置（见类文档），只有
        # "转正"会让某一行从索引里摘掉，不需要整表重扫。
        self._pending_index: dict[tuple[str, str], list[int]] = {}
        self._build_pending_index()
        # 表格最后一行如果是"合计"行（采购单号/型号都是空的，靠公式统计上面的数据区间），
        # 新插入的"已发货"记录要插在它上面，不能插在它下面——见类文档。不是合计行的话就是
        # None，新记录直接接在表格末尾就行，插入都不用。
        self._tail_row = self._detect_tail_row()

    def _detect_tail_row(self) -> int | None:
        r = self._max_row
        if r <= self.header_row:
            return None
        order_no = self.ws.cell(row=r, column=self.col["采购单号"]).value
        model = self.ws.cell(row=r, column=self.col["型号"]).value
        if order_no is None and model is None:
            return r
        return None

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

    def pending_rows(self, order_no: str, model: str) -> list[int]:
        """按行顺序返回这个采购单号+型号名下所有"发货时间=待定 且 状态=未发货"的行号
        （可能不止一个）。直接查 __init__ 时建好的索引，不用每次都扫表；待定行不会挪位置，
        这里返回的行号永远是真实、当前有效的行号。
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
        一行也可能要扣好几行（先到先扣），返回按处理顺序排列的改动列表。直接写，立刻生效，
        不用像插在待定行正上方那套做法一样得先攒一批最后统一处理。
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
            # 被前一轮拆小过，read_quantity 会看到最新的值（待定行是直接写的，立刻生效）。
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

    def _consume_row(
        self, pending_row: int, quantity: int, order_no: str, model: str, zd: str, ship_date: dt.date
    ) -> ShipmentSummaryChange:
        """从这一行待定库存里扣 quantity（调用方保证 0 < quantity <= 这一行当前数量）。"""
        pending_qty = self.read_quantity(pending_row)
        box_capacity = self.ws.cell(row=pending_row, column=self.col["箱容"]).value
        boxes, boxes_exact = _compute_boxes(quantity, box_capacity)

        remaining = pending_qty - quantity
        if remaining > 0:
            # 拆分：待定行原地不动，只把数量/箱数往下调；已发货的这部分单独插一行，放在表格
            # 最下面，不挨着这条待定行。
            remaining_boxes, _ = _compute_boxes(remaining, box_capacity)
            self.ws.cell(row=pending_row, column=self.col["数量"]).value = remaining
            # remaining_boxes 箱容缺失/不合法时会是 None——用 .value = 显式赋值，不然
            # ws.cell(..., value=None) 是个 no-op，待定行会留着拆分前的旧箱数，跟拆分后
            # 变小的「数量」对不上。
            self.ws.cell(row=pending_row, column=self.col["箱数"]).value = remaining_boxes

            new_row = self._insert_shipped_row(pending_row, order_no, model, quantity, boxes, zd, ship_date)

            return ShipmentSummaryChange(
                kind="insert_above",
                pending_row=pending_row,
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

        # quantity == pending_qty：正好扣完这一行，原地转正，不用另外插行
        self._set_explicit_fields(pending_row, order_no, model, quantity, boxes, zd, ship_date, NEW_STATUS)
        self._blank_fields(pending_row)
        # 这一行发货时间从"待定"变成了具体日期，不再是待定库存，从索引里摘掉，不然后面
        # 同一个采购单号+型号再来一笔分摊，会把这一行当成还能扣的库存重复用。
        self._remove_from_pending_index(order_no, model, pending_row)
        return ShipmentSummaryChange(
            kind="convert_in_place",
            pending_row=pending_row,
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

    def _insert_shipped_row(
        self, template_row: int, order_no: str, model: str, quantity: int, boxes, zd: str, ship_date: dt.date
    ) -> int:
        """在表格最下面插一行"已发货"记录（合计行上面，如果有的话；不然直接接在表格末尾），
        产品相关的静态字段（长宽高/毛重/产品名称……）照抄 template_row（也就是被扣的那条
        待定行）——那是这个 SKU 的产品属性，公式列里"引用自己这一行"的部分改成指向新插入的
        这一行。因为新行永远插在表格最下面，需要挪动的最多只有"合计行"这一行本身，代价是
        常数，跟待定行在表格哪个位置、这一批发货有多少笔分摊都没关系。
        """
        if self._tail_row is not None:
            new_row = self._tail_row
            self._push_tail_row_down(1)
        else:
            new_row = self._max_row + 1
            self._max_row = new_row

        self._copy_row_with_reindex(new_row, template_row)
        self._set_explicit_fields(new_row, order_no, model, quantity, boxes, zd, ship_date, NEW_STATUS)
        self._blank_fields(new_row)
        return new_row

    def _push_tail_row_down(self, amount: int) -> None:
        """把合计行搬到再往下 amount 行的位置，给新插入的记录腾出位置；顺带把合计公式的
        区间结束边界往后扩，让新记录也算进合计里（见 _grow_tail_formula）。只碰合计行这
        一行本身，不像 openpyxl 自带的 insert_rows()——它不管插入点在哪，内部都要把整张表
        所有单元格的坐标先排一次序，对着两三万行的表这一步本身就很贵，这里完全绕开它。
        """
        old_row = self._tail_row
        new_row = old_row + amount
        for c in range(1, self._max_column + 1):
            src_cell = self.ws.cell(row=old_row, column=c)
            value = src_cell.value
            if isinstance(value, str) and value.startswith("="):
                value = _grow_tail_formula(value, old_row, new_row)
            dest_cell = self.ws.cell(row=new_row, column=c)
            dest_cell.value = value
            if src_cell.has_style:
                # 见 _copy_row_with_reindex 里的说明：直接复制 _style 这个小整数数组，
                # 不要经过 cell.font/cell.fill 这些高层属性。
                dest_cell._style = copy(src_cell._style)
        dim = self.ws.row_dimensions.get(old_row)
        if dim is not None:
            dest_dim = self.ws.row_dimensions[new_row]
            dest_dim.height = dim.height
            dest_dim.hidden = dim.hidden
            dest_dim.outlineLevel = dim.outlineLevel
            dest_dim.collapsed = dim.collapsed
        self._tail_row = new_row
        self._max_row = new_row

    def _copy_row_with_reindex(self, dest_row: int, src_row: int) -> None:
        # 整行复制：格式（字体/填充/边框/对齐/数字格式）也要一起抄，不然新插入的行会是一整行
        # 没有任何样式的空白格子，跟旁边的行观感完全不一样。
        for c in range(1, self._max_column + 1):
            src_cell = self.ws.cell(row=src_row, column=c)
            dest_cell = self.ws.cell(row=dest_row, column=c)
            value = src_cell.value
            if isinstance(value, str) and value.startswith("="):
                value = _reindex_formula(value, src_row, dest_row)
            dest_cell.value = value
            if src_cell.has_style:
                # 不要通过 cell.font/cell.fill/... 这些高层属性去读/抄样式——每个单元格实际
                # 存的只是 cell._style 这个 9 个整数的 StyleArray（字体/填充/边框……分别是
                # 指向共享样式表的一个下标），cell.font 这类属性每次访问都会现查一遍共享表、
                # 包一层 StyleProxy 返回；对着 StyleProxy 再 copy() 更是很贵的操作（内部会走
                # to_tree/from_tree 序列化一圈）。直接复制 _style 这个小整数数组，是 openpyxl
                # 自己在 WorksheetCopy 里搬一整个工作表时用的同一种手法（见
                # openpyxl/worksheet/copier.py），既快又是官方认可的搬运方式；用 copy() 是因为
                # StyleArray 这几个下标不能多个单元格共享同一个数组实例——后面谁改了自己的
                # 字体，会连带把共享这个数组的其它单元格也改了。
                dest_cell._style = copy(src_cell._style)
        src_dim = self.ws.row_dimensions.get(src_row)
        if src_dim is not None:
            dest_dim = self.ws.row_dimensions[dest_row]
            dest_dim.height = src_dim.height
            dest_dim.hidden = src_dim.hidden
            dest_dim.outlineLevel = src_dim.outlineLevel
            dest_dim.collapsed = src_dim.collapsed

    def _set_explicit_fields(
        self, row: int, order_no: str, model: str, quantity: int, boxes, zd: str, ship_date: dt.date, status: str
    ) -> None:
        # 统一用 .value = 显式赋值，不用 ws.cell(..., value=X)——那种写法碰到 X 是 None 时
        # （比如箱容缺失、_compute_boxes 算不出箱数）什么都不会写，格子会留着模板行的旧箱数，
        # 跟新的「数量」对不上。
        self.ws.cell(row=row, column=self.col["采购单号"]).value = order_no
        self.ws.cell(row=row, column=self.col["型号"]).value = model
        self.ws.cell(row=row, column=self.col["数量"]).value = quantity
        self.ws.cell(row=row, column=self.col["箱数"]).value = boxes
        self.ws.cell(row=row, column=self.col["ZD"]).value = zd
        self.ws.cell(row=row, column=self.col["发货时间"]).value = dt.datetime.combine(ship_date, dt.time())
        self.ws.cell(row=row, column=self.col["状态"]).value = status

    def _blank_fields(self, row: int) -> None:
        for field in BLANK_FIELDS:
            col = self.col.get(field)
            if col is not None:
                # ws.cell(..., value=None) 在 openpyxl 里不会真的清空格子——cell() 只有 value
                # 不是 None 才会赋值，传 None 等于没传。真实表里的待定行经常已经带着"编号"
                # "备注"这些字段的历史值（比如"无库存9"），这一行转正/拆分出新行的时候必须
                # 显式设 .value 才能真的清掉，不然这些历史备注会原样留在新的已发货/未发货记录上。
                self.ws.cell(row=row, column=col).value = None


def _compute_boxes(quantity: int, box_capacity) -> tuple[float | None, bool]:
    if not isinstance(box_capacity, (int, float)) or box_capacity <= 0:
        return None, False
    boxes = quantity / box_capacity
    exact = float(boxes).is_integer()
    return (int(boxes) if exact else boxes), exact
