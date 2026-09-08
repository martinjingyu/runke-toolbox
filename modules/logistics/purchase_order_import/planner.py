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
from dataclasses import dataclass, field
from pathlib import Path

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
    notes: list[str] = field(default_factory=list)


@dataclass
class SkippedOrder:
    order_no: str
    source_file: str
    reason: str


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
) -> Plan:
    purchase_book = PurchaseBook(purchase_ws)
    existing_order_nos = {r.order_no for r in purchase_book.rows}

    summary_book = ShipmentSummaryBook(summary_ws)
    summary_cols = column_index_map(summary_ws, summary_book.header_row)
    require_columns(summary_cols, ["型号", "工厂", "箱容", "长", "宽", "高", "毛重"], "发货计划汇总表")
    dims_index = _build_dims_index(summary_ws, summary_book.header_row, summary_cols)

    items: list[PlanItem] = []
    skipped_orders: list[SkippedOrder] = []
    skipped_files: list[str] = []
    seen_order_nos: set[str] = set()

    for path in list_order_files(folder):
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

        for line in order.lines:
            dims = dims_index.get((supplier_code, line.model)) if supplier_code else None
            box_capacity = dims.box_capacity if dims else None
            length = dims.length if dims else None
            width = dims.width if dims else None
            height = dims.height if dims else None
            gross_weight = dims.gross_weight if dims else None
            boxes, boxes_exact = _compute_boxes(line.quantity, box_capacity)

            item_notes: list[str] = []
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
    """长/宽/高这几列，固定按第一条数据行（见 _dim_source_row）这一列的表达式，重新指向
    新行；第一行这一列如果本身不是公式（比如刚建表还没来得及套），才退回写死数字。
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


def apply_plan(plan: Plan, purchase_ws: Worksheet, summary_ws: Worksheet) -> None:
    if not plan.items:
        return

    purchase_book = PurchaseBook(purchase_ws)
    p_cols = column_index_map(purchase_ws, purchase_book.header_row)
    require_columns(p_cols, ["序号", "订单号", "采购日期", "交货日期", "供应商名称", "型号", "产品名称", "订单数量", "数量单位"], "采购订单汇总表")
    p_last_row = max((r.row_index for r in purchase_book.rows), default=purchase_book.header_row)
    seq = _next_seq(purchase_ws, purchase_book.header_row, p_cols["序号"], p_last_row)

    # 同一个订单号在采购汇总表里只占一个序号——用 plan.items 里"同订单号的行连在一起"这个
    # 保证（build_plan 按订单一份一份处理，一个订单的所有型号行紧挨着追加），订单号变了才
    # 递增序号，行号range最后统一做单元格合并。
    item_seqs: list[int] = []
    current_seq = seq - 1
    prev_order_no: object = object()
    for item in plan.items:
        if item.order_no != prev_order_no:
            current_seq += 1
            prev_order_no = item.order_no
        item_seqs.append(current_seq)

    summary_book = ShipmentSummaryBook(summary_ws)
    s_cols = column_index_map(summary_ws, summary_book.header_row)
    require_columns(
        s_cols,
        ["采购单号", "型号", "产品名称", "箱数", "箱容", "长", "宽", "高", "毛重", "交货时间", "发货时间", "工厂", "状态"],
        "发货计划汇总表",
    )
    s_last_row = _last_data_row(summary_ws, summary_book.header_row, s_cols["采购单号"])

    p_template_row = p_last_row
    s_template_row = s_last_row

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
        p_row = p_last_row + 1 + i
        copy_row(purchase_ws, dest_row=p_row, src_row=p_template_row)
        for c in p_batch_cols:
            _set(purchase_ws, p_row, c, None)
        if p_remark_col:
            _set(purchase_ws, p_row, p_remark_col, None)

        # 序号只写在同一订单号的第一行，后面几行留空，等下面统一合并单元格——跟纸质表格
        # 里"同一个订单号只写一次序号、其余行合并"的排版方式一致。
        is_first_of_order = i == 0 or plan.items[i - 1].order_no != item.order_no
        if is_first_of_order:
            _set(purchase_ws, p_row, p_cols["序号"], str(item_seqs[i]).zfill(_SEQ_WIDTH))
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
        copy_row(summary_ws, dest_row=s_row, src_row=s_template_row)
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
        # "长/宽/高"这几列，真实表里出现过同一个表头名字重复出现不止一次的情况（历史遗留的
        # 重复列，比如后来又加了一组通过 XLOOKUP 从产品信息表查出来的"长/宽/高"）——s_cols
        # 只留了第一次出现的位置，这里改成对每一组出现都单独处理，不能只处理第一组。三个
        # 名字按左到右的顺序一一配对成组（zip），因为"长/宽/高"是一起出现的，同一组的三个
        # 数字应该一起处理（尤其是一整组由同一个数组公式算出来的情况，见 _apply_dim_group）。
        length_cols = all_columns_named(summary_ws, summary_book.header_row, "长")
        width_cols = all_columns_named(summary_ws, summary_book.header_row, "宽")
        height_cols = all_columns_named(summary_ws, summary_book.header_row, "高")
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

    _merge_order_seq_cells(purchase_ws, plan.items, p_cols["序号"], p_last_row)


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
