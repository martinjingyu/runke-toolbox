"""读取发货计划表，只找「标签」在 wanted_labels 里的、「仓库」包含 warehouse_code、
「状态=未发货」、且「发货时间」等于调用方指定的 ship_date 的行——发货时间由用户在界面上选，
不是写死"待定"或者"只要没发货就算数"：同一个仓库、同一个厂商在发货计划表里完全可能同时挂着
好几个不同日期的待发记录（有的已经排了具体发货日期、有的还在"待定"），只有跟这次要处理的这
批标签 PDF/箱唛 PDF 对应的那个发货日期，数字才是对的，不加这个条件的话箱数会把其它日期的
待发记录也一起加进来。

「仓库」按包含匹配（warehouse_code 见 splitter.py 里从 PDF 文件名推导站点代号的说明），不做
精确相等——真实表里「仓库」的值形如"US(CA1)"，不是精确等于站点代号。

「发货时间」这一列在表里读出来是 datetime（比如 datetime(2026, 1, 7, 0, 0)）或者字符串
"待定"这样的特殊值，只有前者能拿来跟 ship_date（一个 date）比较；用 .date() 取出日期部分
比较，不比较时间——真实数据里这一列的时间部分一直是 0 点，但没必要假设永远如此，比较日期
比比较完整 datetime 更稳妥。字符串"待定"这类值直接比较结果是 False，天然被排除，不用额外
判断类型。

wanted_labels 由调用方传入，是标签 PDF 里实际出现过的那些标签——发货计划表本身有几万行，
横跨所有产品全部历史批次，跟这一份标签 PDF 完全无关的标签不该在这里被读出来，也不该被拿去
跟 PDF 做差集提示（不然会报出一大堆"计划里有、这份 PDF 没有"的噪音，那些标签根本不属于
这份 PDF，提示了也没意义）——先看这份标签 PDF 里到底有哪些标签，再拿着这份名单去计划表里逐个
核对，这才是这个工具该有的方向。

按「标签」分组、把这些行的「工厂」和「箱数」汇总起来，给拆分标签 PDF 用：
    标签 -> 这一批待发记录的工厂 + 箱数总和

同一个标签匹配到的待发行「工厂」不一致的话不硬选一个——这种情况单独记下来（见
PendingGroups.conflicts），交给调用方决定怎么提示人工核对，不在这里报错中断整个读取。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from ..shipment_plan_apply.column_utils import column_index_map, find_header_row, require_columns

REQUIRED_HEADERS = ["标签", "状态", "箱数", "工厂", "仓库", "发货时间"]
REQUIRED_HEADERS_BY_FACTORY = ["状态", "箱数", "工厂", "仓库", "发货时间"]

PENDING_STATUS = "未发货"


@dataclass
class PendingGroup:
    label: str
    factory: str
    total_boxes: float
    boxes_exact: bool
    row_count: int


@dataclass
class PendingGroups:
    groups: dict[str, PendingGroup] = field(default_factory=dict)
    conflicts: dict[str, list[str]] = field(default_factory=dict)  # 标签 -> 不一致的工厂列表


def _num(value) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    raise ValueError(f"「箱数」列的值「{value}」不是数字")


def _matches_ship_date(value, ship_date: dt.date) -> bool:
    if isinstance(value, dt.datetime):
        return value.date() == ship_date
    if isinstance(value, dt.date):
        return value == ship_date
    return False


def load_pending_groups(
    xlsx_path: str | Path,
    wanted_labels: set[str],
    warehouse_code: str,
    ship_date: dt.date,
    sheet_name: str | None = None,
) -> PendingGroups:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    header_row = find_header_row(ws, REQUIRED_HEADERS, max_scan_rows=10)
    cols = column_index_map(ws, header_row)
    idx = require_columns(cols, REQUIRED_HEADERS, "发货计划表")

    factories: dict[str, set[str]] = {}
    boxes: dict[str, float] = {}
    counts: dict[str, int] = {}

    for row in ws.iter_rows(min_row=header_row + 1):
        label = row[idx["标签"] - 1].value
        if label is None:
            continue
        label = str(label).strip()
        if label not in wanted_labels:
            continue

        status = row[idx["状态"] - 1].value
        if status != PENDING_STATUS:
            continue

        if not _matches_ship_date(row[idx["发货时间"] - 1].value, ship_date):
            continue

        warehouse = row[idx["仓库"] - 1].value
        if warehouse is None or warehouse_code not in str(warehouse):
            continue

        factory = row[idx["工厂"] - 1].value
        factory = str(factory).strip() if factory is not None else ""
        boxes_value = row[idx["箱数"] - 1].value

        factories.setdefault(label, set()).add(factory)
        boxes[label] = boxes.get(label, 0.0) + _num(boxes_value)
        counts[label] = counts.get(label, 0) + 1

    result = PendingGroups()
    for label, fs in factories.items():
        if len(fs) > 1:
            result.conflicts[label] = sorted(fs)
            continue
        total = boxes[label]
        exact = float(total).is_integer()
        result.groups[label] = PendingGroup(
            label=label,
            factory=next(iter(fs)),
            total_boxes=int(total) if exact else total,
            boxes_exact=exact,
            row_count=counts[label],
        )
    return result


def load_pending_boxes_by_factory(
    xlsx_path: str | Path, warehouse_code: str, ship_date: dt.date, sheet_name: str | None = None
) -> dict[str, float]:
    """跟 load_pending_groups 一样按「仓库含 warehouse_code + 状态=未发货 + 发货时间=
    ship_date」筛选，但不看「标签」，直接按「工厂」把「箱数」加总——给 lowm_splitter.py 用：
    LO-WM 站点的箱唛 PDF 不需要认每一页具体是哪个 SKU（Walmart 只在乎这一批货的总箱数，不在乎
    每箱具体装的是什么，见 lowm_splitter.py 顶部说明），只要知道"这个厂商这次一共要发多少箱"
    就够了。
    """
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    header_row = find_header_row(ws, REQUIRED_HEADERS_BY_FACTORY, max_scan_rows=10)
    cols = column_index_map(ws, header_row)
    idx = require_columns(cols, REQUIRED_HEADERS_BY_FACTORY, "发货计划表")

    totals: dict[str, float] = {}
    for row in ws.iter_rows(min_row=header_row + 1):
        status = row[idx["状态"] - 1].value
        if status != PENDING_STATUS:
            continue

        if not _matches_ship_date(row[idx["发货时间"] - 1].value, ship_date):
            continue

        warehouse = row[idx["仓库"] - 1].value
        if warehouse is None or warehouse_code not in str(warehouse):
            continue

        factory = row[idx["工厂"] - 1].value
        factory = str(factory).strip() if factory is not None else ""
        boxes_value = row[idx["箱数"] - 1].value

        totals[factory] = totals.get(factory, 0.0) + _num(boxes_value)
    return totals
