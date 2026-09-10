"""把解析、货号翻译、采购分摊、发货计划汇总表更新串起来，分成"算一遍"和"真的写"两步：

- build_plan()：只读、只在内存里模拟分摊结果，完全不碰 openpyxl 的写操作，不管这一批发货
  计划有没有问题都能跑完、把所有问题一次性收集出来。
- apply_plan()：只有在 build_plan() 判定"整批没有任何错误"的前提下才能调用，这时候才真的
  往采购汇总表、发货计划汇总表的工作表对象里写数据（插入/转正发货计划行；日期列必须已经
  存在，不会现场插入，见 build_plan 里对日期列的检查）。

这样保证"分摊到一半才发现后面数量不够"不会导致文件被写了一半——build_plan 阶段发现任何
问题，整批直接不进入 apply_plan，跟之前和业务确认过的"要么整批成功，要么什么都不改"一致。

即使 apply_plan() 跑完，也只是改了内存里的 openpyxl Workbook 对象，实际磁盘上的文件要调用方
自己在人工确认之后再 wb.save()——这个模块不负责存盘，也不负责备份。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from .product_lookup import ProductLookup
from .purchase_book import Allocation, PurchaseBook
from .shipment_summary import ShipmentSummaryBook, ShipmentSummaryChange
from .shipment_templates import PlanLine


@dataclass
class PlanItem:
    line: PlanLine
    huohao: str | None = None
    allocations: list[Allocation] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


@dataclass
class Plan:
    ship_date: dt.date
    items: list[PlanItem]
    parse_errors: list[str]

    @property
    def has_blocking_errors(self) -> bool:
        return bool(self.parse_errors) or any(item.errors for item in self.items)

    @property
    def total_allocations(self) -> int:
        return sum(len(item.allocations) for item in self.items)


def _line_label(line: PlanLine) -> str:
    prefix = f"[{line.source_file}] " if line.source_file else ""
    return f"{prefix}第{line.source_row}行"


def build_plan(
    lines: list[PlanLine],
    parse_errors: list[str],
    lookup: ProductLookup,
    purchase_book: PurchaseBook,
    ship_date: dt.date,
) -> Plan:
    items: list[PlanItem] = []

    # 日期列必须已经在采购订单汇总表里手动加好——早期版本会在这里现场插入一新列，但插入会
    # 挪动它右边所有列的位置，实际使用中出现过插入点右边的公式跟着错位的问题，现在改成只读
    # 查找，找不到就在这里报出来，不进入 apply 阶段去写（见 purchase_book.py 顶部说明）。
    if purchase_book.find_date_column(ship_date) is None:
        parse_errors = list(parse_errors) + [
            f"采购订单汇总表里还没有 {ship_date.strftime('%Y-%m-%d')} 这一天的日期列——"
            "请先在表格里手动加好这一天的日期列，再重新更新"
        ]
        return Plan(ship_date=ship_date, items=items, parse_errors=parse_errors)

    for line in lines:
        item = PlanItem(line=line)
        huohao = lookup.resolve(line.sku_kind, line.sku)
        if huohao is None:
            item.errors.append(
                f"{_line_label(line)}：SKU「{line.sku}」在在售产品信息总表里查不到对应的货号（同款）"
            )
            items.append(item)
            continue

        item.huohao = huohao
        outcome = purchase_book.allocate(huohao, line.quantity)
        item.allocations = outcome.allocations

        if outcome.shortfall > 0:
            item.errors.append(
                f"{_line_label(line)}：货号「{huohao}」相关采购订单的未出货数量"
                f"加起来还差 {outcome.shortfall} 个，凑不够这次要发的 {line.quantity} 个"
            )

        items.append(item)

    return Plan(
        ship_date=ship_date,
        items=items,
        parse_errors=list(parse_errors),
    )


def build_plan_from_recorded_allocations(
    lines: list[PlanLine],
    parse_errors: list[str],
    lookup: ProductLookup,
    purchase_book: PurchaseBook,
    ship_date: dt.date,
) -> Plan:
    """"发货计划汇总表更新"专用——不按余量现算分摊，而是读采购表里 ship_date 那一天已经
    写了的量（「采购订单分摊更新」写的），按同样的先后顺序重放出"这条发货计划该冲抵哪笔
    采购订单"，只用来确定发货计划汇总表待定行该匹配哪个（采购单号, 型号），不再重新校验
    余量够不够——那一步判断在写采购表的时候已经做过了，这里传进来的采购表就是已经用过的，
    重复校验没有意义（见跟用户的讨论）。如果读出来的记录量对不上这批计划要的量，说明传
    进来的采购表可能跟这批发货计划不是同一批，仍然会报错，但报的是"记录对不上"而不是
    "余量不够"。
    """
    items: list[PlanItem] = []

    date_col = purchase_book.find_date_column(ship_date)
    if date_col is None:
        parse_errors = list(parse_errors) + [
            f"采购订单汇总表里还没有 {ship_date.strftime('%Y-%m-%d')} 这一天的记录——"
            "请先用「采购订单分摊更新」把这一批写进采购订单汇总表，再用这个工具"
        ]
        return Plan(ship_date=ship_date, items=items, parse_errors=parse_errors)

    for line in lines:
        item = PlanItem(line=line)
        huohao = lookup.resolve(line.sku_kind, line.sku)
        if huohao is None:
            item.errors.append(
                f"{_line_label(line)}：SKU「{line.sku}」在在售产品信息总表里查不到对应的货号（同款）"
            )
            items.append(item)
            continue

        item.huohao = huohao
        outcome = purchase_book.allocate_recorded(huohao, date_col, line.quantity)
        item.allocations = outcome.allocations

        if outcome.shortfall > 0:
            item.errors.append(
                f"{_line_label(line)}：货号「{huohao}」在采购订单汇总表 "
                f"{ship_date.strftime('%Y-%m-%d')} 这一列记录的量加起来还差 {outcome.shortfall} 个，"
                f"凑不够这次要发的 {line.quantity} 个——可能这份发货计划表跟当初"
                "「采购订单分摊更新」用的不是同一批，请核对"
            )

        items.append(item)

    return Plan(
        ship_date=ship_date,
        items=items,
        parse_errors=list(parse_errors),
    )


def apply_plan(
    plan: Plan,
    purchase_book: PurchaseBook,
    summary_book: ShipmentSummaryBook,
    progress_callback=None,
) -> list[ShipmentSummaryChange]:
    # progress_callback(done, total)：按"已经处理了几笔分摊"报进度，不是只有个转圈圈的忙碌条。
    # summary_book.apply_shipment 现在是直接写、立刻生效——新插入的"已发货"记录固定放在表格
    # 最下面（不挨着被扣的待定行），代价只跟"插了几条新记录"有关，跟待定行在表格哪个位置、
    # 表格本身多大都没关系，所以不用像之前那样先攒一批改动最后再统一处理（见
    # shipment_summary.py 顶部说明）。
    if plan.has_blocking_errors:
        raise ValueError("这一批发货计划里还有没解决的错误，不能写入")

    date_col = purchase_book.require_date_column(plan.ship_date)

    total = plan.total_allocations
    done = 0
    changes: list[ShipmentSummaryChange] = []
    for item in plan.items:
        for allocation in item.allocations:
            purchase_book.write_allocation(allocation, date_col)
            # 发货计划汇总表里同一个采购单号+型号可能有不止一条待定行（都是同一批还没决定去
            # 哪的库存，行与行之间没有另外的区别）——apply_shipment 会按行顺序依次扣，一次
            # 分摊可能因此产生不止一条改动，见 shipment_summary.py 的说明。
            new_changes = summary_book.apply_shipment(
                order_no=allocation.row.order_no,
                model=allocation.row.model,
                quantity=allocation.quantity,
                zd=item.line.zd,
                ship_date=plan.ship_date,
                sku=item.line.sku,
            )
            changes.extend(new_changes)
            done += 1
            if progress_callback is not None:
                progress_callback(done, total)

    # 这一批新插入的记录，Excel 的筛选范围不会自动跟着扩大——不补这一步的话，数据其实写对了，
    # 但在 Excel 里拿筛选框去找刚写的记录会找不到，容易被误以为没写进去（见
    # shipment_summary.py 里 sync_auto_filter 的说明）。
    summary_book.sync_auto_filter()

    return changes


def apply_plan_purchase_only(
    plan: Plan,
    purchase_book: PurchaseBook,
    progress_callback=None,
) -> None:
    """"采购订单分摊更新"这个工具用——只往采购订单汇总表里写分摊，完全不碰发货计划汇总表
    （那张表按 ZD 拆目的地是"发货计划自动更新"的活，这边不需要）。

    同一个货号（型号）在这一批里可能因为分属不同的目的地（不同 ZD）被分成好几条 PlanItem，
    但采购订单汇总表的日期列本来就是"这一天一共发了多少"的汇总，不分 ZD——`write_allocation`
    早就是"在已有数量上累加"而不是覆盖（见 purchase_book.py），所以同一个货号不同 ZD 的
    好几笔分摊，写到同一个日期列里自然就是叠加在一起，不用在这里另外处理。
    """
    if plan.has_blocking_errors:
        raise ValueError("这一批发货计划里还有没解决的错误，不能写入")

    date_col = purchase_book.require_date_column(plan.ship_date)

    total = plan.total_allocations
    done = 0
    for item in plan.items:
        for allocation in item.allocations:
            purchase_book.write_allocation(allocation, date_col)
            done += 1
            if progress_callback is not None:
                progress_callback(done, total)


def apply_plan_summary_only(
    plan: Plan,
    summary_book: ShipmentSummaryBook,
    progress_callback=None,
) -> list[ShipmentSummaryChange]:
    """"发货计划汇总表更新"这个工具用——只往发货计划汇总表里写待发货记录，完全不碰采购订单
    汇总表（不插日期列、不写累计出货量，那是"采购订单分摊更新"/"发货计划自动更新"的活）。

    注意：build_plan() 用来算"这个货号该分摊到哪几笔采购订单"的 purchase_book 还是要传，
    因为发货计划汇总表里待定行的匹配 key 是 (采购单号, 型号)——这两个字段来自 purchase_book
    分摊出来的 Allocation，不是随便传的；只是这里只读它算出来的结果，不会往它里面写任何东西、
    也不需要保存它。调用方（面板）不该给这张表做备份/存盘，省了这一步风险和耗时。
    """
    if plan.has_blocking_errors:
        raise ValueError("这一批发货计划里还有没解决的错误，不能写入")

    total = plan.total_allocations
    done = 0
    changes: list[ShipmentSummaryChange] = []
    for item in plan.items:
        for allocation in item.allocations:
            new_changes = summary_book.apply_shipment(
                order_no=allocation.row.order_no,
                model=allocation.row.model,
                quantity=allocation.quantity,
                zd=item.line.zd,
                ship_date=plan.ship_date,
                sku=item.line.sku,
            )
            changes.extend(new_changes)
            done += 1
            if progress_callback is not None:
                progress_callback(done, total)

    summary_book.sync_auto_filter()

    return changes
