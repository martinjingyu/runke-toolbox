"""把"文件夹里的一批采购订单文件"变成"采购汇总表/发货计划汇总表各要新增哪些行"——分两步：

- build_plan()：只读，扫描文件夹、按订单号去重、给每个型号算供应商代码/箱数/长宽高毛重，
  不碰 openpyxl 的写操作。任何字段缺信息（供应商没映射、型号没有同工厂历史记录、箱数除不尽）
  都不是硬错误——直接留空，把情况记进这一行的 notes 里，跟 build_plan 结果一起摆给人看，
  不阻断其它行/其它订单继续处理（这一点跟 shipment_plan_apply 的"整批要么全过要么全不写"
  不一样：那边改的是已有库存记录，错一条会让数字对不上，这边是从零新增一整行，新增本身
  没有"对不上"的风险，缺信息也不影响别的行，所以做成尽量填、填不出来就留空+提示）。
- apply_plan()：在 build_plan() 的基础上，真的往两个 workbook 对象里追加新行（只在内存，
  不存盘，调用方自己决定什么时候 wb.save()）。两张表都是"接到表格现有数据最后一行继续往下追加"，
  不需要像 shipment_summary.py 那样在表格中间 insert_rows——那是"从已有的一条待定库存里分一部分
  出去"，这里是"这个型号在表里还完全没出现过"，直接接在最后一行下面写，不会打乱任何公式的
  行号引用（采购汇总表里 SUBTOTAL 的区间本来就写死到很大的行号，发货计划汇总表本身没有这种
  跨行公式，都不受追加影响）。

  新行不是凭空写一堆空格子——先用 column_utils.copy_row() 把"现有最后一行"整行复制过去
  （格式、公式都跟着抄一份，公式里"引用自己这一行"的部分会自动指向新行，见 copy_row 的
  说明），保证新行看起来跟旁边的行一样、公式列（店铺/标签/DP/CBM/未出货数量等）不用业务
  人员自己再拖一遍。复制完之后，再把这一行"每条订单各不相同"的字段（型号/数量/供应商代码
  等业务值）和"这一行是全新库存、不该继承模板行历史状态"的字段（采购汇总表的出货批次数量列、
  发货计划汇总表的仓库/FBA ID/追踪编号/ZD/编号/备注/货代/出货单号）覆盖成正确的值——前者
  覆盖成这一条记录自己的数据，后者统一覆盖成空，不然会把模板行"已经发生过的历史"错当成
  新订单的状态抄过来。

两张表要求的字段来源、留空的字段，见模块设计文档；这里只放实现。
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from ..shipment_plan_apply.column_utils import (
    all_columns_named,
    column_index_map,
    copy_row,
    is_formula_value,
    reindex_formula_value,
    require_columns,
    unmerge_overlapping_rows,
)
from ..shipment_plan_apply.purchase_book import PurchaseBook
from ..shipment_plan_apply.shipment_summary import BLANK_FIELDS, NEW_STATUS, PENDING_LABEL, ShipmentSummaryBook
from .order_file import OrderFile, OrderFileError, list_order_files, parse_order_file

_UNIT = "pcs"
_SEQ_WIDTH = 3

# 「序号」这一列真实表里不是这张表自己按行独立编的号，是订单号自带的信息——订单号形如
# "SX-2609215"（工厂-年月+三位流水号），最后三位数字"215"直接就是这一单该填的「序号」，
# 跟订单文件名前缀（"215.采购订单SX-2609215.xlsx"）对得上，业务方本来就是照着订单号手填的
# （核对过真实表格，"TM-2503045"对应序号"045"，"WJ-2609213"对应序号"213"，一个不差）。
_ORDER_SEQ_RE = re.compile(r"(\d{3})$")


def _parse_seq_from_order_no(order_no: str) -> str | None:
    m = _ORDER_SEQ_RE.search(order_no)
    return m.group(1) if m else None


# 发货计划汇总表「箱数/数量/CBM/总材重/总实重」这几列，业务要求在表头正上方那一行（跟采购
# 汇总表第 1 行放 SUBTOTAL 的位置是同一个道理——header_row 再往上一行）放一个"这一列往下
# 全部数据的合计"，方便打开表格一眼看到总数，不用自己拉到最后一行或者手动选中一整列。
# 区间结束行号写死成一个很大的数（不是"当前表格实际最后一行"），这样新增行只要没超过这个
# 数，公式本身完全不用跟着改——跟采购汇总表那边 SUBTOTAL(9,I4:I1000732) 是同一个思路，
# 见 planner.py 顶部说明。每次批量导入完都重新写一遍这几个公式（覆盖写，不是只在缺失时才
# 补）：这样万一被人手动改坏、删掉过，下次导入会自动纠正回来，不用另外记着去修。
_SUM_RANGE_END_ROW = 1_000_000
_SUM_COLUMNS = ["箱数", "数量", "CBM", "总材重", "总实重"]


def _write_summary_column_sums(ws: Worksheet, header_row: int, cols: dict[str, int]) -> None:
    total_row = header_row - 1
    first_data_row = header_row + 1
    for name in _SUM_COLUMNS:
        col = cols[name]
        letter = get_column_letter(col)
        formula = f"=SUBTOTAL(9,{letter}{first_data_row}:{letter}{_SUM_RANGE_END_ROW})"
        ws.cell(row=total_row, column=col).value = formula


@dataclass
class HistoricalDims:
    box_capacity: float | None
    length: float | None
    width: float | None
    height: float | None
    gross_weight: float | None
    source_row: int


@dataclass
class PlanItem:
    order_no: str
    model: str
    product_name: str
    quantity: int
    purchase_date: dt.date | None
    delivery_date: dt.date | None
    supplier_name: str | None
    supplier_code: str | None
    box_capacity: float | None
    length: float | None
    width: float | None
    height: float | None
    gross_weight: float | None
    boxes: float | None
    boxes_exact: bool
    source_file: str
    source_row: int
    seq_hint: str | None  # 从订单号解析出来的「序号」，见 _parse_seq_from_order_no；解析不出来才是 None
    notes: list[str] = field(default_factory=list)


@dataclass
class SkippedOrder:
    order_no: str
    source_file: str
    reason: str


class PlannerError(Exception):
    """apply_plan 发现即将写入的位置跟预期的不一样（比如那个位置本来就有内容），
    没法安全地继续，调用方应该把这个错误原样展示给人看。"""


@dataclass
class Plan:
    items: list[PlanItem]
    skipped_orders: list[SkippedOrder]
    skipped_files: list[str]  # 整份文件解析失败（比如提取不出订单号），文件名+原因已经拼进字符串里

    @property
    def total_lines(self) -> int:
        return len(self.items)


def _num(value) -> float | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return float(value)


def _compute_boxes(quantity: int, box_capacity: float | None) -> tuple[float | None, bool]:
    if not box_capacity or box_capacity <= 0:
        return None, False
    boxes = quantity / box_capacity
    exact = float(boxes).is_integer()
    return (int(boxes) if exact else boxes), exact


def _build_dims_index(ws: Worksheet, header_row: int, col: dict[str, int]) -> dict[tuple[str, str], HistoricalDims]:
    """按(工厂, 型号)建索引：同一个组合出现好几次的话，行号更大（更晚追加）的覆盖前面的，
    最终每个组合留下的是"最近一次"的箱容/长宽高/毛重。
    """
    c_model = col["型号"]
    c_factory = col["工厂"]
    c_capacity = col["箱容"]
    c_length = col["长"]
    c_width = col["宽"]
    c_height = col["高"]
    c_weight = col["毛重"]
    max_col = max(c_model, c_factory, c_capacity, c_length, c_width, c_height, c_weight)

    index: dict[tuple[str, str], HistoricalDims] = {}
    for row_no, row in enumerate(
        ws.iter_rows(min_row=header_row + 1, max_col=max_col, values_only=True), start=header_row + 1
    ):
        model = row[c_model - 1]
        factory = row[c_factory - 1]
        if not model or not factory:
            continue
        key = (str(factory).strip(), str(model).strip())
        index[key] = HistoricalDims(
            box_capacity=_num(row[c_capacity - 1]),
            length=_num(row[c_length - 1]),
            width=_num(row[c_width - 1]),
            height=_num(row[c_height - 1]),
            gross_weight=_num(row[c_weight - 1]),
            source_row=row_no,
        )
    return index


def _last_data_row(ws: Worksheet, header_row: int, key_col: int) -> int:
    """往下扫，找最后一行 key_col 有值的行号——表格实际数据结尾往往比 ws.max_row 小很多
    （尾部一大片只是带样式的空行），不能直接拿 ws.max_row 当追加位置。
    """
    last = header_row
    for row_no, value in enumerate(
        ws.iter_rows(min_row=header_row + 1, max_col=key_col, values_only=True), start=header_row + 1
    ):
        if value[key_col - 1] is not None:
            last = row_no
    return last


def _assert_rows_blank(ws: Worksheet, start_row: int, end_row: int, max_col: int, context: str) -> None:
    """真的往下写之前，确认 [start_row, end_row] 这一段目前确实是空的——"最后一行"是靠扫描
    某一列（订单号/型号）找出来的，前提是这一列在真正的数据行里从来不为空、在这一列之后
    只剩空白行。真实表格里踩过的坑：表格末尾偶尔会有一行"合计"/统计公式（订单号/型号这两列
    刚好留空，但别的列还有内容），前面找"最后一行"的逻辑会把这种行当成"空白"跳过，
    误以为要写入的位置是空的，实际上一写就会把这一行的公式/内容整个覆盖掉，而且覆盖后
    看起来还是正常的一行数据，很难事后发现。这里改成写之前先整段扫一遍确认真的没有内容，
    一旦发现不是空的就直接停下来报错，交给人去看表格结构，不能猜着往下写。
    """
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, max_col=max_col):
        for cell in row:
            if cell.value is not None:
                raise PlannerError(
                    f"{context}：准备追加新记录的位置——第 {start_row} 到 {end_row} 行——"
                    f"里，第 {cell.row} 行第 {cell.column} 列已经有内容（{cell.value!r}），"
                    f"不像是空行。可能是「表格最后一行在哪」判断错了（比如那一行其实是合计/"
                    f"统计公式，不是真正的订单数据），为了不覆盖已有内容，这里先停下来，"
                    f"需要人工检查一下这张表格的结构"
                )


def _next_seq(ws: Worksheet, header_row: int, seq_col: int, last_row: int) -> int:
    max_seq = 0
    for row in ws.iter_rows(min_row=header_row + 1, max_row=last_row, min_col=seq_col, max_col=seq_col, values_only=True):
        try:
            max_seq = max(max_seq, int(str(row[0])))
        except (TypeError, ValueError):
            continue
    return max_seq + 1


def build_plan(
    folder: Path,
    purchase_ws: Worksheet,
    summary_ws: Worksheet,
    supplier_map: dict[str, str],
    progress_callback=None,
) -> Plan:
    """progress_callback(stage: str, done: int, total: int)：给界面上报进度用——真实的采购
    汇总表/发货计划汇总表都是几千到几万行，PurchaseBook/ShipmentSummaryBook 本身已经支持按
    批次上报加载进度（见 purchase_book.py/shipment_summary.py），之前这里没接这根线，界面上
    只能看到一个转圈圈的"正在处理"，看不出究竟卡在哪一步、还要多久。跟
    shipment_plan_apply/purchase_only_panel.py 里 _PreviewWorker 的用法保持一致：不传的话
    什么都不做，兼容只关心结果、不需要进度的调用方（比如测试）。
    """

    def _report(stage: str, done: int = 0, total: int = 0) -> None:
        if progress_callback is not None:
            progress_callback(stage, done, total)

    _report("正在加载采购订单汇总表")
    purchase_book = PurchaseBook(
        purchase_ws, progress_callback=lambda done, total: _report("正在加载采购订单汇总表", done, total)
    )
    existing_order_nos = {r.order_no for r in purchase_book.rows}

    _report("正在加载发货计划汇总表")
    summary_book = ShipmentSummaryBook(
        summary_ws, progress_callback=lambda done, total: _report("正在加载发货计划汇总表", done, total)
    )
    summary_cols = column_index_map(summary_ws, summary_book.header_row)
    require_columns(summary_cols, ["型号", "工厂", "箱容", "长", "宽", "高", "毛重"], "发货计划汇总表")
    dims_index = _build_dims_index(summary_ws, summary_book.header_row, summary_cols)

    items: list[PlanItem] = []
    skipped_orders: list[SkippedOrder] = []
    skipped_files: list[str] = []
    seen_order_nos: set[str] = set()

    order_paths = list_order_files(folder)
    for file_i, path in enumerate(order_paths, start=1):
        _report("正在解析订单文件", file_i, len(order_paths))
        try:
            order = parse_order_file(path)
        except OrderFileError as exc:
            skipped_files.append(f"「{path.name}」：{exc}")
            continue

        if order.order_no in existing_order_nos or order.order_no in seen_order_nos:
            skipped_orders.append(
                SkippedOrder(order_no=order.order_no, source_file=path.name, reason="采购汇总表里已经存在这个订单号")
            )
            continue
        seen_order_nos.add(order.order_no)

        supplier_code = supplier_map.get(order.supplier_name.strip()) if order.supplier_name else None
        seq_hint = _parse_seq_from_order_no(order.order_no)

        for line in order.lines:
            dims = dims_index.get((supplier_code, line.model)) if supplier_code else None
            box_capacity = dims.box_capacity if dims else None
            length = dims.length if dims else None
            width = dims.width if dims else None
            height = dims.height if dims else None
            gross_weight = dims.gross_weight if dims else None
            boxes, boxes_exact = _compute_boxes(line.quantity, box_capacity)

            item_notes: list[str] = []
            if seq_hint is None:
                item_notes.append(
                    f"订单号「{order.order_no}」不是「XX-YYMMNNN」这种格式，解析不出最后三位当「序号」，"
                    "写入时会退回按表格里已有的最大序号顺延，请核对写进去的序号对不对"
                )
            if order.supplier_name is None:
                item_notes.append("订单文件里没找到供应商信息，「供应商名称」「工厂」留空")
            elif supplier_code is None:
                item_notes.append(f"供应商「{order.supplier_name}」还没配置映射代码，「供应商名称」「工厂」留空")
            if supplier_code is not None and dims is None:
                item_notes.append(f"型号「{line.model}」在供应商「{supplier_code}」名下没有历史记录，箱容/长/宽/高/毛重留空")
            if order.purchase_date is None:
                item_notes.append("订单文件里没找到采购日期，留空")
            if boxes is not None and not boxes_exact:
                item_notes.append(f"订单数量 {line.quantity} 除以箱容 {box_capacity} 除不尽，箱数留了小数 {boxes}，需要人工核对")
            if line.delivery_date is None:
                item_notes.append("这一行没读到交货日期，留空")

            items.append(
                PlanItem(
                    order_no=order.order_no,
                    model=line.model,
                    product_name=line.product_name,
                    quantity=line.quantity,
                    purchase_date=order.purchase_date,
                    delivery_date=line.delivery_date,
                    supplier_name=order.supplier_name,
                    supplier_code=supplier_code,
                    box_capacity=box_capacity,
                    length=length,
                    width=width,
                    height=height,
                    gross_weight=gross_weight,
                    boxes=boxes,
                    boxes_exact=boxes_exact,
                    source_file=path.name,
                    source_row=line.source_row,
                    seq_hint=seq_hint,
                    notes=item_notes,
                )
            )

    return Plan(items=items, skipped_orders=skipped_orders, skipped_files=skipped_files)


def _as_datetime(value: dt.date | None) -> dt.datetime | None:
    return dt.datetime.combine(value, dt.time()) if value is not None else None


def _dim_source_row(header_row: int) -> int:
    # 真实表里出现过"长/宽/高"这几列历史上没有统一套用公式——早期数据是手填的写死数字，
    # 中间某次改版才开始套公式，且中途偶尔还会混进个别残缺/不完整的异常行（比如公式对不上
    # 别的行）。与其"以当前表格最后一行为准、找不到再往上找最近一行"（一旦最近的那一行刚好
    # 是异常数据，会悄悄拷贝出错误结果，很难发现），不如固定认表格第一条数据行——业务方只要
    # 保证这一行的公式/表达式是对的，新增行就总能拿到正确、统一的结果，不用管中间这段历史
    # 数据有多混乱。
    return header_row + 1


def _apply_dim_column(ws: Worksheet, col: int, header_row: int, dest_row: int, fallback: float | None) -> None:
    """固定按第一条数据行（见 _dim_source_row）这一列的表达式，重新指向新行；第一行这一列
    如果本身不是公式（比如刚建表还没来得及套），才退回 fallback。

    最早只是给"长/宽/高"用的，后来发现"数量"（=箱数*箱容）、"DP"（按标签 VLOOKUP 查外部
    工作簿）这两列真实表里是同一种历史包袱：第一条数据行是公式，但表格现在的最后一行（也就是
    新增行原本要抄的模板行）不一定还是公式——可能是历史上某次手动改成写死数字/字符串，或者
    这一行本身就属于那批"早年没套公式"的老数据。如果照抄模板行，新增行会把这一行当时凑巧写
    死的旧数字/旧文本原样继承过来，跟这一行真正对应的型号/供应商完全对不上（实测："数量"
    抄到了模板行历史遗留的旧数字，"DP"抄到了模板行所在供应商的旧代码，新增的其它供应商的
    型号全部被错误地打上了同一个 DP）。所以这两列现在也复用这个函数，而不是让 copy_row()
    整行复制的默认行为决定。
    """
    source_row = _dim_source_row(header_row)
    value = ws.cell(row=source_row, column=col).value
    if not is_formula_value(value):
        _set(ws, dest_row, col, fallback)
        return
    _set(ws, dest_row, col, reindex_formula_value(value, source_row, dest_row))


def _apply_dim_group(
    ws: Worksheet,
    length_col: int,
    width_col: int,
    height_col: int,
    header_row: int,
    dest_row: int,
    length: float | None,
    width: float | None,
    height: float | None,
) -> None:
    """真实表里出现过"长"是一个数组公式（比如一次 XLOOKUP 查出长宽高三个值，公式原文只存
    在"长"这一格，"宽"/"高"在文件里读出来只是缓存的普通数字，永远不会有公式文本，见
    column_utils.is_formula_value 的说明）——这种情况下"宽"/"高"不能各自独立按
    _apply_dim_column 处理（不管找哪一行都不会有属于自己的公式，只会一直退回写死数字）。
    只要第一条数据行这一格解出来是数组公式，"宽"/"高"就是它的溢出结果，不用（也不该）单独
    写值——Excel 打开重新计算的时候会根据这个数组公式自动算出正确结果；这里主动清空，避免
    残留旧行的、不会再更新的旧数字长期显示在表里误导人。
    """
    source_row = _dim_source_row(header_row)
    value = ws.cell(row=source_row, column=length_col).value
    if isinstance(value, ArrayFormula):
        _set(ws, dest_row, length_col, reindex_formula_value(value, source_row, dest_row))
        _set(ws, dest_row, width_col, None)
        _set(ws, dest_row, height_col, None)
        return
    # 不是数组公式（普通字符串公式/写死数字/完全没有）——退回原来"三列各自独立处理"的策略
    _apply_dim_column(ws, length_col, header_row, dest_row, length)
    _apply_dim_column(ws, width_col, header_row, dest_row, width)
    _apply_dim_column(ws, height_col, header_row, dest_row, height)


def _set(ws: Worksheet, row: int, col: int, value) -> None:
    # ws.cell(row, column, value=X) 在 openpyxl 里只有 X 不是 None 才会真的赋值——传 None
    # 等于没传，格子会保留 copy_row() 抄过来的旧值。这里统一走 .value = ，None 也能真的清空。
    ws.cell(row=row, column=col).value = value


def apply_plan(plan: Plan, purchase_ws: Worksheet, summary_ws: Worksheet, progress_callback=None) -> None:
    """progress_callback(done: int, total: int)：给界面上报"正在写入第几条"用——采购汇总表
    那张 sheet 的 ws.max_column 实测被撑到了 16384（某个很远的格子只是留了点格式，见
    column_utils.copy_row 的说明），整行复制一次要碰这么多格子，批量导入几十上百条订单的话
    这一步本身就不便宜，不接进度的话界面在这期间只能干等。
    """
    if not plan.items:
        return

    purchase_book = PurchaseBook(purchase_ws)
    p_cols = column_index_map(purchase_ws, purchase_book.header_row)
    require_columns(p_cols, ["序号", "订单号", "采购日期", "交货日期", "供应商名称", "型号", "产品名称", "订单数量", "数量单位"], "采购订单汇总表")
    p_last_row = max((r.row_index for r in purchase_book.rows), default=purchase_book.header_row)

    # 同一个订单号在采购汇总表里只占一个序号——用 plan.items 里"同订单号的行连在一起"这个
    # 保证（build_plan 按订单一份一份处理，一个订单的所有型号行紧挨着追加），订单号变了才
    # 换下一个序号，行号range最后统一做单元格合并。
    #
    # 序号本身优先直接用 item.seq_hint（从订单号自己解析出来的，见 _parse_seq_from_order_no
    # 的说明）——这张表真实数据验证过，「序号」就是订单号最后三位数字，不是这张表自己按行
    # 独立编的号，甚至不是按行号单调递增的：这批表格历史数据里出现过更早的行序号反而比更晚
    # 的行更大（序号的计数周期在中间某处——大概率是年份——重新起过），拿"全表历史最大值+1"
    # 会算出一个跟订单号本身完全对不上、还比合理值大出一截的号。只有 seq_hint 解析不出来
    # （订单号格式不是"XX-YYMMNNN"）才退回旧的"表格里已有的最大序号+1，同一批里继续往后顺延"
    # 这个退路——而且只在真用到的时候才去扫一遍表格算最大值，正常情况（订单号都能解析出
    # 序号）不需要付这个扫描代价。
    _fallback_next: int | None = None

    def _next_fallback_seq() -> str:
        nonlocal _fallback_next
        if _fallback_next is None:
            _fallback_next = _next_seq(purchase_ws, purchase_book.header_row, p_cols["序号"], p_last_row)
        seq = _fallback_next
        _fallback_next += 1
        return str(seq).zfill(_SEQ_WIDTH)

    item_seq_strs: list[str] = []
    current_seq_str: str | None = None
    prev_order_no: object = object()
    for item in plan.items:
        if item.order_no != prev_order_no:
            current_seq_str = item.seq_hint if item.seq_hint is not None else _next_fallback_seq()
            prev_order_no = item.order_no
        item_seq_strs.append(current_seq_str)

    summary_book = ShipmentSummaryBook(summary_ws)
    s_cols = column_index_map(summary_ws, summary_book.header_row)
    require_columns(
        s_cols,
        [
            "采购单号", "型号", "产品名称", "箱数", "箱容", "数量", "长", "宽", "高", "毛重",
            "交货时间", "发货时间", "工厂", "状态", "CBM", "总材重", "总实重",
        ],
        "发货计划汇总表",
    )
    s_last_row = _last_data_row(summary_ws, summary_book.header_row, s_cols["采购单号"])

    p_template_row = p_last_row
    s_template_row = s_last_row

    # "长/宽/高"这三个表头在哪些列出现过（真实表里出现过不止一组，见 _apply_dim_group 的
    # 说明）只跟表头这一行有关，每一条新记录都是同一个答案——放到循环外面算一次，不然每新增
    # 一行都要把整张表头重新扫一遍（all_columns_named 内部是 ws[header_row]，代价跟下面
    # ws.max_column 的问题是一回事，只是这里连"扫出来的答案"都没变过，完全没必要每行重复扫）。
    length_cols = all_columns_named(summary_ws, summary_book.header_row, "长")
    width_cols = all_columns_named(summary_ws, summary_book.header_row, "宽")
    height_cols = all_columns_named(summary_ws, summary_book.header_row, "高")

    # 两张表实际有数据的列范围——不能拿 ws.max_column 当数，真实采购汇总表这张属性被撑到过
    # 16384（某个很远的格子只是留了点格式，不代表真的有这么多列有数据，见 column_utils.py
    # 里 unmerge_overlapping_columns 附近的说明）。这个范围既用来确认要写入的位置真的是空的
    # （_assert_rows_blank），也传给下面的 copy_row()——copy_row 不传 max_col 的话会退回
    # ws.max_column，整行复制这一步就要陪着这 16384 列空跑一遍；实测这一项差异是「167ms/行」
    # 还是「36ms/行」的区别，导入的订单越多，这个不必要的常数因子被放大的次数就越多。
    #
    # 发货计划汇总表这边不能只看 s_cols（column_index_map 同名表头只留第一次出现的位置，见
    # 它的说明）——"长/宽/高"真实表里出现过不止一组，第二组要靠 length_cols/width_cols/
    # height_cols 才找得到，直接并进来取最大值，不能只留一个拍脑袋的安全余量指望"差不多够"。
    p_max_col = purchase_book.remaining_col + 5
    s_max_col = max([*s_cols.values(), *length_cols, *width_cols, *height_cols], default=0) + 5

    _assert_rows_blank(purchase_ws, p_last_row + 1, p_last_row + len(plan.items), p_max_col, "采购订单汇总表")
    _assert_rows_blank(summary_ws, s_last_row + 1, s_last_row + len(plan.items), s_max_col, "发货计划汇总表")

    # 真实表格里常见手工把某一列"提前"合并了一大片还没用到的空白行（比如序号列一路合并到
    # 很靠后的行）——即将追加的新行如果正好落进这种旧合并区域，openpyxl 会把对应格子当成
    # MergedCell，连值都赋不了。这两张表接下来各要新增 len(plan.items) 行，先把这个范围内
    # 任何跟它有重叠的旧合并区域拆开，保证下面 copy_row()/_set() 能正常写值；不影响这个
    # 范围之外的合并（比如已有数据里正常的、同一订单多型号共用一个序号的合并）。
    unmerge_overlapping_rows(purchase_ws, p_last_row + 1, p_last_row + len(plan.items))
    unmerge_overlapping_rows(summary_ws, s_last_row + 1, s_last_row + len(plan.items))

    # 采购汇总表「数量单位」和「未出货数量」之间那一大片是各批次的已出货数量——这些是模板行
    # 自己的出货历史，新订单还没发过货，这些格子照抄过来的话会凭空多出一堆假的已出货记录，
    # 必须显式清空（未出货数量本身是公式，留着不动，会在 Excel 里根据清空后的批次列自动算对）。
    p_batch_cols = range(purchase_book.date_col_start, purchase_book.date_col_end + 1)
    p_remark_col = p_cols.get("备注")

    for i, item in enumerate(plan.items):
        if progress_callback is not None:
            progress_callback(i, len(plan.items))
        p_row = p_last_row + 1 + i
        copy_row(purchase_ws, dest_row=p_row, src_row=p_template_row, max_col=p_max_col)
        for c in p_batch_cols:
            _set(purchase_ws, p_row, c, None)
        if p_remark_col:
            _set(purchase_ws, p_row, p_remark_col, None)

        # 序号只写在同一订单号的第一行，后面几行留空，等下面统一合并单元格——跟纸质表格
        # 里"同一个订单号只写一次序号、其余行合并"的排版方式一致。
        is_first_of_order = i == 0 or plan.items[i - 1].order_no != item.order_no
        if is_first_of_order:
            _set(purchase_ws, p_row, p_cols["序号"], item_seq_strs[i])
        else:
            _set(purchase_ws, p_row, p_cols["序号"], None)
        _set(purchase_ws, p_row, p_cols["订单号"], item.order_no)
        _set(purchase_ws, p_row, p_cols["采购日期"], _as_datetime(item.purchase_date))
        _set(purchase_ws, p_row, p_cols["交货日期"], _as_datetime(item.delivery_date))
        _set(purchase_ws, p_row, p_cols["供应商名称"], item.supplier_code)
        _set(purchase_ws, p_row, p_cols["型号"], item.model)
        _set(purchase_ws, p_row, p_cols["产品名称"], item.product_name)
        _set(purchase_ws, p_row, p_cols["订单数量"], item.quantity)
        _set(purchase_ws, p_row, p_cols["数量单位"], _UNIT)

        s_row = s_last_row + 1 + i
        copy_row(summary_ws, dest_row=s_row, src_row=s_template_row, max_col=s_max_col)
        # 仓库/FBA ID/追踪编号/编号/备注/货代/出货单号（BLANK_FIELDS）+ ZD：都是模板行"这一笔
        # 具体是怎么发出去的"记录，新订单还没分配到任何一次具体发货，照抄过来就是张冠李戴。
        for name in [*BLANK_FIELDS, "ZD"]:
            col = s_cols.get(name)
            if col:
                _set(summary_ws, s_row, col, None)

        _set(summary_ws, s_row, s_cols["采购单号"], item.order_no)
        _set(summary_ws, s_row, s_cols["型号"], item.model)
        _set(summary_ws, s_row, s_cols["产品名称"], item.product_name)
        _set(summary_ws, s_row, s_cols["箱数"], item.boxes)
        _set(summary_ws, s_row, s_cols["箱容"], item.box_capacity)
        # "数量"固定按第一条数据行的表达式（一般是"=箱数*箱容"）重新指向新行，不直接照抄
        # copy_row() 从模板行带过来的值——见 _apply_dim_column 的说明，这一列同样存在"模板行
        # 恰好是历史上写死数字的那种行，抄过来就是错的旧数字"这个问题，fallback 用这一条记录
        # 自己算出来的订单数量（item.quantity 就是订单原始数量，箱数*箱容凑不满整数箱的时候
        # 未必等于它，但比模板行的旧数字更接近事实）。
        _apply_dim_column(summary_ws, s_cols["数量"], summary_book.header_row, s_row, item.quantity)
        # "DP"不是每张表都一定有（不是必须列），有的话同样按第一条数据行的表达式重新指向新行；
        # 没有可用表达式（且这一列存在）的话留空，不猜——不能直接照抄模板行，见上面的说明。
        dp_col = s_cols.get("DP")
        if dp_col:
            _apply_dim_column(summary_ws, dp_col, summary_book.header_row, s_row, None)
        # "长/宽/高"这几列，真实表里出现过同一个表头名字重复出现不止一次的情况（历史遗留的
        # 重复列，比如后来又加了一组通过 XLOOKUP 从产品信息表查出来的"长/宽/高"）——s_cols
        # 只留了第一次出现的位置，这里改成对每一组出现都单独处理，不能只处理第一组。三个
        # 名字按左到右的顺序一一配对成组（zip），因为"长/宽/高"是一起出现的，同一组的三个
        # 数字应该一起处理（尤其是一整组由同一个数组公式算出来的情况，见 _apply_dim_group）。
        # length_cols/width_cols/height_cols 在循环外面已经算好了（哪些列叫这几个名字只跟
        # 表头有关，每条记录都是同一个答案，不用每行重新扫一遍表头）。
        for length_col, width_col, height_col in zip(length_cols, width_cols, height_cols):
            _apply_dim_group(
                summary_ws,
                length_col,
                width_col,
                height_col,
                summary_book.header_row,
                s_row,
                item.length,
                item.width,
                item.height,
            )
        _set(summary_ws, s_row, s_cols["毛重"], item.gross_weight)
        _set(summary_ws, s_row, s_cols["交货时间"], _as_datetime(item.delivery_date))
        _set(summary_ws, s_row, s_cols["发货时间"], PENDING_LABEL)
        _set(summary_ws, s_row, s_cols["工厂"], item.supplier_code)
        _set(summary_ws, s_row, s_cols["状态"], NEW_STATUS)

    if progress_callback is not None:
        progress_callback(len(plan.items), len(plan.items))

    _merge_order_seq_cells(purchase_ws, plan.items, p_cols["序号"], p_last_row)
    _write_summary_column_sums(summary_ws, summary_book.header_row, s_cols)


def _merge_order_seq_cells(ws: Worksheet, items: list[PlanItem], seq_col: int, first_new_row: int) -> None:
    """同一订单号新增的几行连在一起，把它们的「序号」列合并成一个单元格（值已经只写在
    第一行，其余行留空，交给合并单元格来显示成"同一个序号跨了好几行"）。
    """
    run_start = first_new_row + 1
    for i in range(1, len(items) + 1):
        row = first_new_row + i
        is_last = i == len(items) or items[i].order_no != items[i - 1].order_no
        if is_last:
            if row > run_start:
                ws.merge_cells(start_row=run_start, start_column=seq_col, end_row=row, end_column=seq_col)
            run_start = row + 1
