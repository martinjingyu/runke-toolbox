"""把解析、货号翻译、采购分摊、发货计划汇总表更新串起来，分成"算一遍"和"真的写"两步：

- build_plan()：只读、只在内存里模拟分摊结果，完全不碰 openpyxl 的写操作，不管这一批发货
  计划有没有问题都能跑完、把所有问题一次性收集出来。
- apply_plan()：只有在 build_plan() 判定"整批没有任何错误"的前提下才能调用，这时候才真的
  往采购汇总表、发货计划汇总表的工作表对象里写数据（插入/转正发货计划行；日期列必须已经
  存在，不会现场插入，见 build_plan 里对日期列的检查）。

build_plan 阶段发现的问题分两种，处理方式不一样：SKU 查不到货号、解析报错这些是数据本身
有问题，仍然会让整批直接不进入 apply_plan，跟之前和业务确认过的"要么整批成功，要么什么都
不改"一致（见 Plan.has_blocking_errors）。"余量不够"是唯一的例外——只把这一行本身跳过
（不分摊、不写），记进 PlanItem.skip_reason，不影响同一批里其它行照常写入；被跳过的行
用 write_skipped_items_report() 单独汇总成一张异常数据表格，见 build_plan 里
_TEMPLATE_PRIORITY 附近的说明。

即使 apply_plan() 跑完，也只是改了内存里的 openpyxl Workbook 对象，实际磁盘上的文件要调用方
自己在人工确认之后再 wb.save()——这个模块不负责存盘，也不负责备份。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

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
    # 余量不够时不再算成阻塞整批的 errors，而是跳过这一行（allocations 留空、什么都不写），
    # 原因记在这里——跟 errors 分开是因为这两种情况现在处理方式完全不同：errors 非空会让
    # Plan.has_blocking_errors 变 True、整批都不写；skip_reason 非空只是这一行自己不写，
    # 其它行照常处理，见 build_plan 里 _TEMPLATE_PRIORITY 附近的说明。
    skip_reason: str | None = None


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

    @property
    def skipped_items(self) -> list[PlanItem]:
        return [item for item in self.items if item.skip_reason]


def _line_label(line: PlanLine) -> str:
    prefix = f"[{line.source_file}] " if line.source_file else ""
    return f"{prefix}第{line.source_row}行"


# 处理顺序：亚马逊 > 沃尔玛 > 海外仓——库存是几个平台共用的同一批采购订单，先到先得（见
# purchase_book.allocate 的说明），谁先分摊就更容易分到货。业务上海外仓补货没有平台账号
# 被压评分/限制上架这种硬约束，缺货可以晚几天再发，亚马逊/沃尔玛断货的代价更大，所以特意
# 把海外仓排在最后——真缺货的话，缺口优先落在海外仓身上，不是随便哪个平台撞上就算倒霉。
# 没有 template_type（理论上不会发生，parse_shipment_plan 统一写过）的排在最后、跟海外仓
# 同一优先级，不让它意外插到亚马逊/沃尔玛前面抢库存。
_TEMPLATE_PRIORITY = {"amazon": 0, "walmart": 1, "overseas": 2}


def _template_priority(line: PlanLine) -> int:
    return _TEMPLATE_PRIORITY.get(line.template_type, len(_TEMPLATE_PRIORITY))


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

    # 稳定排序：同一模板内部原来的先后顺序不变，只是把不同模板的行分组、按优先级重新排列。
    ordered_lines = sorted(lines, key=_template_priority)

    for line in ordered_lines:
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

        if outcome.shortfall > 0:
            # 跳过、不写——不算整批的阻塞错误。回滚这次分摊已经占用的余量，不然这个注定要
            # 跳过的分摊会白白占掉后面同货号其它行本来还能分到的库存（见 PurchaseRow.remaining
            # 的算法：outcome.allocations 里的每一份在 purchase_book.allocate 内部已经把
            # row_obj.consumed_this_run 加过了，这里必须原样减回去）。
            for allocation in outcome.allocations:
                allocation.row.consumed_this_run -= allocation.quantity
            item.skip_reason = (
                f"{_line_label(line)}：货号「{huohao}」相关采购订单的未出货数量"
                f"加起来还差 {outcome.shortfall} 个，凑不够这次要发的 {line.quantity} 个，已跳过"
            )
        else:
            item.allocations = outcome.allocations

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
    余量够不够——那一步判断在写采购表的时候已经做过了。

    这里同一条发货计划表传的还是原始输入（运营的发货计划表本身不会因为「采购订单分摊更新"
    跳过了某一行就自动消失），"记录的量不够"现在分不清到底是两种情况里的哪一种：① 这一行
    当初就是因为余量不够被「采购订单分摊更新」跳过的（正常情况，见 build_plan 的说明）；
    ② 传进来的采购表其实跟这批发货计划不是同一批（数据真的对不上）。既然分不清，就按跟
    build_plan 一样的规则处理——跳过、不算整批的阻塞错误，把原因记到 item.skip_reason，
    调用方会把这些行也汇总进异常数据表格（见 write_skipped_items_report）；真是传错了
    采购表这种情况，异常表格里一整批都会是"记录不够"，人一眼就能看出来是文件传错了，不需要
    靠这里硬报错去提醒。
    """
    items: list[PlanItem] = []

    date_col = purchase_book.find_date_column(ship_date)
    if date_col is None:
        parse_errors = list(parse_errors) + [
            f"采购订单汇总表里还没有 {ship_date.strftime('%Y-%m-%d')} 这一天的记录——"
            "请先用「采购订单分摊更新」把这一批写进采购订单汇总表，再用这个工具"
        ]
        return Plan(ship_date=ship_date, items=items, parse_errors=parse_errors)

    # 跟 build_plan 一样按亚马逊>沃尔玛>海外仓重排——道理是一样的：如果这一批记录的量本身
    # 就不够摊给所有行（不管是当初写入时就不够、还是这之后又有改动），谁先重放到就先分到，
    # 跟原始写入时用的是同一套优先级，回放结果才跟"当初到底发生了什么"保持一致。
    ordered_lines = sorted(lines, key=_template_priority)

    for line in ordered_lines:
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

        if outcome.shortfall > 0:
            # 见上面 docstring：分不清是"当初就被跳过"还是"传错了采购表"，统一按跳过处理，
            # 回滚这次重放已经占用的记录量，不然会白白占掉后面同货号其它行本来能对上的量。
            for allocation in outcome.allocations:
                allocation.row.consumed_this_run -= allocation.quantity
            item.skip_reason = (
                f"{_line_label(line)}：货号「{huohao}」在采购订单汇总表 "
                f"{ship_date.strftime('%Y-%m-%d')} 这一列记录的量加起来还差 {outcome.shortfall} 个，"
                f"凑不够这次要发的 {line.quantity} 个（可能当初就是因为余量不够被跳过了，也可能"
                "这份发货计划表跟当初「采购订单分摊更新」用的不是同一批），已跳过"
            )
        else:
            item.allocations = outcome.allocations

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


_TEMPLATE_LABELS_ZH = {"amazon": "亚马逊", "walmart": "沃尔玛", "overseas": "海外仓"}


def write_skipped_items_report(skipped_items: list[PlanItem], first_plan_path: Path) -> Path | None:
    """把因为余量不够被跳过的发货计划行汇总成一张"异常数据表格"，存在这一批第一份发货计划表
    旁边（同一个文件夹）——方便运营回头看这一批里到底哪些没发出去、该补哪些货。

    没有任何被跳过的行就不生成文件，不然每次正常跑完都平白多出一个空表，反而让人分不清
    "这次真的有问题"还是"这个工具每次都会生成一个"。文件名带时间戳，避免同一天跑了好几批、
    后一次的异常表把前一次的覆盖掉——时间戳精确到秒，真要是同一秒内跑了不止一次（比如手滑
    连点，或者「采购订单分摊更新」跟「发货计划汇总表更新」紧挨着跑、用的是同一份发货计划表），
    只精确到秒还是会撞上、静默覆盖掉前一份，所以撞了名字的话再加个序号后缀，直到找到一个
    没被占用的文件名，不覆盖任何已有文件。
    """
    if not skipped_items:
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "异常数据"
    ws.append(["来源文件", "表格行数", "平台", "SKU 种类", "SKU", "货号", "目的地/ZD", "需要数量", "说明"])
    for item in skipped_items:
        line = item.line
        ws.append(
            [
                line.source_file,
                line.source_row,
                _TEMPLATE_LABELS_ZH.get(line.template_type, line.template_type or ""),
                line.sku_kind,
                line.sku,
                item.huohao or "",
                line.destination_label,
                line.quantity,
                item.skip_reason or "",
            ]
        )
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 18

    timestamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    base_name = f"{first_plan_path.stem}-异常数据-{timestamp}"
    out_path = first_plan_path.parent / f"{base_name}.xlsx"
    suffix = 2
    while out_path.exists():
        out_path = first_plan_path.parent / f"{base_name}-{suffix}.xlsx"
        suffix += 1
    wb.save(out_path)
    return out_path
