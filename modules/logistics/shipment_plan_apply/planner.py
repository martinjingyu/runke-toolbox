"""把解析、货号翻译、采购分摊、发货计划汇总表更新串起来，分成"算一遍"和"真的写"两步：

- build_plan()：只读、只在内存里模拟分摊结果，完全不碰 openpyxl 的写操作，不管这一批发货
  计划有没有问题都能跑完、把所有问题一次性收集出来。
- apply_plan()：只有在 build_plan() 判定"整批没有任何错误"的前提下才能调用，这时候才真的
  往采购汇总表、发货计划汇总表的工作表对象里写数据（包括插入日期列、插入/转正发货计划行）。

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
    used_variant: bool = False
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
        variant = lookup.variant_of(huohao)
        outcome = purchase_book.allocate(huohao, line.quantity, variant)
        item.allocations = outcome.allocations
        item.used_variant = outcome.used_variant

        if outcome.shortfall > 0:
            variant_note = f"（含变体货号「{variant}」）" if variant else ""
            item.errors.append(
                f"{_line_label(line)}：货号「{huohao}」{variant_note}相关采购订单的未出货数量"
                f"加起来还差 {outcome.shortfall} 个，凑不够这次要发的 {line.quantity} 个"
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
    # progress_callback(done, total)：这一步慢的地方几乎全在 summary_book.apply_shipment 里的
    # 插入行操作——发货计划汇总表有两三万行，插一行要把插入点以下所有行的公式都重新扫一遍
    # （见 shipment_summary.py），一批发货计划有几十条分摊记录的话，跑完可能要几分钟，所以
    # 按"已经处理了几笔分摊"报进度，不是只有个转圈圈的忙碌条。
    if plan.has_blocking_errors:
        raise ValueError("这一批发货计划里还有没解决的错误，不能写入")

    date_col = purchase_book.find_or_create_date_column(plan.ship_date)

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
            )
            changes.extend(new_changes)
            done += 1
            if progress_callback is not None:
                progress_callback(done, total)

    return changes
