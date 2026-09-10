import datetime as dt
from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import Border, PatternFill, Side

from core.diff_preview import GROUP_KEY
from modules.logistics.shipment_plan_apply.column_utils import HeaderNotFoundError, resolve_cell_value
from modules.logistics.shipment_plan_apply.diff import (
    run_and_capture_diff,
    run_and_capture_diff_purchase_only,
    run_and_capture_diff_summary_only,
)
from modules.logistics.shipment_plan_apply.planner import (
    build_plan,
    build_plan_from_recorded_allocations,
    apply_plan,
    apply_plan_purchase_only,
    apply_plan_summary_only,
)
from modules.logistics.shipment_plan_apply.product_lookup import ProductLookupError, load_product_lookup
from modules.logistics.shipment_plan_apply.purchase_book import DateColumnNotFoundError, PurchaseBook
from modules.logistics.shipment_plan_apply.shipment_summary import ShipmentSummaryBook
from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine, parse_shipment_plan

REAL_DATA_DIR = Path("/Users/jingyuhuang/Documents/Work/闰科/物流仓库/采购汇总+发货计划")


# ---------------------------------------------------------------------------
# column_utils.resolve_cell_value：预览用的"公式尽量算出结果"逻辑
# ---------------------------------------------------------------------------


def test_resolve_cell_value_evaluates_local_arithmetic_and_concat(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    # A2=采购单号(带字母，模拟真实订单号) B2=型号 C2=箱数 D2=箱容 E2=CBM公式
    # F2=标签(自引用型号) G2=拼接(采购单号+型号，模拟真实的 "&" 那一列)
    ws.append(["采购单号", "型号", "箱数", "箱容", "CBM", "标签", "拼接"])
    ws.append(["GH-2501009", "TD-RZ-419", 3, 12, "=C2*D2", "=+B2", "=A2&B2"])
    wb.save(tmp_path / "f.xlsx")
    wb2 = openpyxl.load_workbook(tmp_path / "f.xlsx", data_only=False)
    ws2 = wb2.active

    assert resolve_cell_value(ws2, 2, 5) == 3 * 12  # CBM 列，纯数字运算
    assert resolve_cell_value(ws2, 2, 6) == "TD-RZ-419"  # 标签 = 自引用型号（字符串也要能算）
    # "&" 拼接：真实数据里两边都是带字母的编号（比如订单号、型号），不是纯数字，之前的实现
    # 会因为替换后的表达式里出现字母被当成"不安全"而算不出来，这里就是回归这个问题
    assert resolve_cell_value(ws2, 2, 7) == "GH-2501009TD-RZ-419"


def test_resolve_cell_value_gives_up_on_external_or_function_formulas(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["A", "B"])
    ws.append(["=VLOOKUP(A1,B:C,2,0)", "=SUM(A1:A5)"])
    wb.save(tmp_path / "f.xlsx")
    wb2 = openpyxl.load_workbook(tmp_path / "f.xlsx", data_only=False)
    ws2 = wb2.active

    assert resolve_cell_value(ws2, 2, 1) is None
    assert resolve_cell_value(ws2, 2, 2) is None


def test_resolve_cell_value_passes_through_plain_values():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([42, "文本", None])
    assert resolve_cell_value(ws, 1, 1) == 42
    assert resolve_cell_value(ws, 1, 2) == "文本"
    assert resolve_cell_value(ws, 1, 3) is None


# ---------------------------------------------------------------------------
# shipment_templates
# ---------------------------------------------------------------------------


def _write_walmart_plan(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["店铺", "RK-SKU", "GTIN", "WM-SKU", "Item name", "数量"])
    ws.append(["CK-沃尔玛", "TD-348", "gtin1", "WM-1", "Table Lamp", 21])
    ws.append([None, "TD-392", "gtin2", "WM-2", "Table Lamp", 0])  # 数量为 0，应该报错
    ws.append([None, "TD-521", "gtin3", "WM-3", "Table Lamp", 30])
    wb.save(path)


def test_parse_walmart_plan(tmp_path):
    path = tmp_path / "walmart.xlsx"
    _write_walmart_plan(path)
    plan = parse_shipment_plan(path, "Sheet")
    assert plan.template_type == "walmart"
    assert len(plan.lines) == 2  # 数量为 0 的那行被判定非法，不计入 lines
    assert len(plan.errors) == 1
    assert "正数" in plan.errors[0]
    # ZD 不是店铺名字本身，是按店铺名字里的关键字映射出来的到站编号
    assert plan.lines[0].zd == "CK-WM"
    assert plan.lines[0].destination_label == "CK-沃尔玛"  # 店铺原始名字留给展示用
    assert plan.lines[0].sku_kind == "RK"
    assert plan.lines[0].sku == "TD-348"
    assert plan.lines[1].zd == "CK-WM"  # 店铺是靠"沿用上一个非空值"填下来的


def test_parse_walmart_plan_maps_lo_shop_to_lo_wm(tmp_path):
    path = tmp_path / "walmart.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["店铺", "RK-SKU", "GTIN", "WM-SKU", "Item name", "数量"])
    ws.append(["SX-LO-沃尔玛", "TD-158", "gtin1", "WM-1", "Table Lamp", 15])
    wb.save(path)

    plan = parse_shipment_plan(path, "Sheet")
    assert not plan.errors
    assert plan.lines[0].zd == "LO-WM"


def test_parse_walmart_plan_unrecognized_shop_name_reports_error(tmp_path):
    path = tmp_path / "walmart.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["店铺", "RK-SKU", "GTIN", "WM-SKU", "Item name", "数量"])
    ws.append(["未知店铺", "TD-158", "gtin1", "WM-1", "Table Lamp", 15])
    wb.save(path)

    plan = parse_shipment_plan(path, "Sheet")
    assert not plan.lines
    assert len(plan.errors) == 1
    assert "未知店铺" in plan.errors[0]


def _write_amazon_plan(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["发货计划", None, 1])
    ws.append(["店铺", "SKU", "US", "CA"])
    ws.append(["cinkeda", "TD-CKD-206", 30, None])
    ws.append([None, "TD-CK-584", None, 21])
    ws.append([None, "TD-CK-57", 21, 21])  # 同一行两个目的地都有数量，要拆成两条
    wb.save(path)


def test_parse_amazon_plan_splits_multi_destination_rows(tmp_path):
    path = tmp_path / "amazon.xlsx"
    _write_amazon_plan(path)
    plan = parse_shipment_plan(path, "Sheet")
    assert plan.template_type == "amazon"
    assert len(plan.lines) == 4
    dest_by_sku = {(l.sku, l.destination_label): l.quantity for l in plan.lines}
    assert dest_by_sku[("TD-CKD-206", "US")] == 30
    assert dest_by_sku[("TD-CK-584", "CA")] == 21
    # ZD 要填的是这一行发去的目的地（US/CA），不是"店铺"（cinkeda）——后面写发货计划汇总表的
    # ZD 列，看的就是这个字段
    assert all(l.zd == l.destination_label for l in plan.lines)
    assert dest_by_sku[("TD-CK-57", "US")] == 21
    assert dest_by_sku[("TD-CK-57", "CA")] == 21
    assert all(l.sku_kind == "AMZ" for l in plan.lines)


def test_parse_amazon_plan_ignores_missing_or_absent_shop_column(tmp_path):
    # 亚马逊表不需要匹配"店铺"——只认 SKU + 站点列 + 数量，"店铺"缺失/整份表压根没有
    # 这一列都不该报错（见 _parse_amazon 的说明，店铺只标注卖家账号，不影响 ZD/库存分摊）。
    path = tmp_path / "amazon_no_shop_value.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["店铺", "SKU", "US", "CA"])
    ws.append([None, "TD-CKD-206", 30, None])  # 这一行、以及它之前都没出现过店铺
    wb.save(path)
    plan = parse_shipment_plan(path, "Sheet")
    assert not plan.errors
    assert len(plan.lines) == 1
    assert plan.lines[0].sku == "TD-CKD-206"
    assert plan.lines[0].zd == "US"

    path2 = tmp_path / "amazon_no_shop_column.xlsx"
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.append(["SKU", "US", "CA"])  # 整份表压根没有"店铺"这一列
    ws2.append(["TD-CKD-206", 30, None])
    wb2.save(path2)
    plan2 = parse_shipment_plan(path2, "Sheet")
    assert plan2.template_type == "amazon"
    assert not plan2.errors
    assert len(plan2.lines) == 1
    assert plan2.lines[0].sku == "TD-CKD-206"
    assert plan2.lines[0].zd == "US"


def _write_overseas_plan(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["图片", "海外仓-SKU", "9.2发货", None])
    ws.append([None, None, "CA1", "WF-RX-MD"])
    ws.append([None, "TD-158", 30, None])
    ws.append([None, "TD-244", 30, 60])  # 两个目的仓都有数量
    wb.save(path)


def test_parse_overseas_plan(tmp_path):
    path = tmp_path / "overseas.xlsx"
    _write_overseas_plan(path)
    plan = parse_shipment_plan(path, "Sheet")
    assert plan.template_type == "overseas"
    assert len(plan.lines) == 3
    dest_by_sku = {(l.sku, l.destination_label): l.quantity for l in plan.lines}
    assert dest_by_sku[("TD-158", "CA1")] == 30
    assert dest_by_sku[("TD-244", "CA1")] == 30
    assert dest_by_sku[("TD-244", "WF-RX-MD")] == 60
    assert all(l.zd == l.destination_label for l in plan.lines)  # 海外仓：ZD 就是目的仓列头


def test_parse_shipment_plan_missing_headers_names_the_file_and_sheet(tmp_path):
    # 回归测试：这个工具一次要读好几张不同的表，报错只说"找不到表头"不说是哪张表/哪个
    # sheet，人没法一眼定位该去检查哪份文件。文件名和 sheet 名都要出现在报错里。
    path = tmp_path / "运营发的沃尔玛计划.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "9月计划"
    ws.append(["店铺", "RK-SKU", "没有数量这一列"])
    ws.append(["CK-沃尔玛", "TD-348", 21])
    wb.save(path)

    with pytest.raises(HeaderNotFoundError) as exc_info:
        parse_shipment_plan(path, "9月计划", template_type="walmart")
    message = str(exc_info.value)
    assert "运营发的沃尔玛计划.xlsx" in message
    assert "9月计划" in message


# ---------------------------------------------------------------------------
# product_lookup
# ---------------------------------------------------------------------------


def _write_product_info(path: Path, rows: list[tuple]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["AMZ-SKU", "RK-SKU", "同款"])
    for row in rows:
        ws.append(list(row))
    wb.save(path)


def test_product_lookup_resolve(tmp_path):
    path = tmp_path / "product.xlsx"
    _write_product_info(
        path,
        [
            ("AMZ-1", "RK-1", "HH-1"),
            ("AMZ-2", "RK-2", "HH-2"),
        ],
    )
    lookup = load_product_lookup(path)
    assert lookup.resolve("AMZ", "AMZ-1") == "HH-1"
    assert lookup.resolve("RK", "RK-2") == "HH-2"
    assert lookup.resolve("RK", "不存在") is None


def test_product_lookup_raises_on_conflicting_mapping(tmp_path):
    path = tmp_path / "product.xlsx"
    _write_product_info(
        path,
        [
            ("AMZ-1", "RK-1", "HH-1"),
            ("AMZ-1", "RK-9", "HH-9"),  # 同一个 AMZ-SKU 映射到两个不同货号
        ],
    )
    with pytest.raises(ProductLookupError):
        load_product_lookup(path)


def test_product_lookup_missing_headers_names_the_table(tmp_path):
    path = tmp_path / "product.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["AMZ-SKU", "RK-SKU"])  # 缺"同款"这一列
    wb.save(path)

    with pytest.raises(HeaderNotFoundError, match="在售产品信息总表"):
        load_product_lookup(path)


# ---------------------------------------------------------------------------
# purchase_book
# ---------------------------------------------------------------------------


def _write_purchase_book(path: Path, include_2026_09_01_col: bool = True):
    # 日期列必须已经存在于表格里才能写（不再现场插入，见 purchase_book.py），所以这里除了
    # 两个历史日期列（2024-01-01/2024-02-01），默认还预先放一个空的 2026-09-01 列——这是
    # 这份测试文件里绝大多数发货计划用的日期，供 require_date_column/build_plan 找到并写入。
    # include_2026_09_01_col=False：专门给"日期列压根不存在"这类场景用，别加这一列。
    wb = openpyxl.Workbook()
    ws = wb.active
    header = ["订单号", "采购日期", "型号", "订单数量", "数量单位", dt.datetime(2024, 1, 1), dt.datetime(2024, 2, 1)]
    sub_header = [None, None, None, None, None, "出货时间", "出货时间"]
    row3 = ["PO-EARLY", dt.datetime(2024, 1, 1), "M1", 100, "pcs", 20, None]
    row4 = ["PO-LATE", dt.datetime(2024, 6, 1), "M1", 50, "pcs", None, None]
    if include_2026_09_01_col:
        header.append(dt.datetime(2026, 9, 1))
        sub_header.append("出货时间")
        row3.append(None)
        row4.append(None)
        remaining_range = "F{r}:H{r}"
    else:
        remaining_range = "F{r}:G{r}"
    header.append("未出货数量")
    sub_header.append(None)
    row3.append(f"=D3-SUM({remaining_range.format(r=3)})")
    row4.append(f"=D4-SUM({remaining_range.format(r=4)})")
    ws.append(header)
    ws.append(sub_header)
    ws.append(row3)
    ws.append(row4)
    wb.save(path)
    return wb


def test_purchase_book_allocates_earliest_order_first(tmp_path):
    path = tmp_path / "purchase.xlsx"
    _write_purchase_book(path)
    wb = openpyxl.load_workbook(path, data_only=False)
    book = PurchaseBook(wb.active)

    orders = book.by_model["M1"]
    assert [o.order_no for o in orders] == ["PO-EARLY", "PO-LATE"]
    assert orders[0].initial_remaining == 80
    assert orders[1].initial_remaining == 50

    outcome = book.allocate("M1", 100)
    assert outcome.shortfall == 0
    assert [(a.row.order_no, a.quantity) for a in outcome.allocations] == [
        ("PO-EARLY", 80),
        ("PO-LATE", 20),
    ]


def test_purchase_book_reports_shortfall(tmp_path):
    path = tmp_path / "purchase.xlsx"
    _write_purchase_book(path)
    wb = openpyxl.load_workbook(path, data_only=False)
    book = PurchaseBook(wb.active)

    outcome = book.allocate("M1", 500)
    assert outcome.shortfall == 500 - 80 - 50


def test_purchase_book_require_date_column_finds_existing_column(tmp_path):
    path = tmp_path / "purchase.xlsx"
    _write_purchase_book(path)
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb.active
    book = PurchaseBook(ws)

    date_col = book.require_date_column(dt.date(2024, 1, 1))  # 已有的日期列 F

    outcome = book.allocate("M1", 10)
    for a in outcome.allocations:
        book.write_allocation(a, date_col)

    wb.save(path)
    wb2 = openpyxl.load_workbook(path, data_only=False)
    book2 = PurchaseBook(wb2.active)
    row2 = book2.by_model["M1"][0]
    assert row2.initial_remaining == 80 - 10  # 重新加载后公式算出来的未出货数量要正确


def test_purchase_book_require_date_column_raises_when_missing(tmp_path):
    # 日期列不会再自动插入——目标日期没有对应的列时必须直接报错，让人先手动把这一天的
    # 日期列加好，不能由代码现场插一列（见模块文档：早期版本会插，但插入点右边的公式/格式
    # 实际用起来还是出现过错位的问题）。
    path = tmp_path / "purchase.xlsx"
    _write_purchase_book(path)
    wb = openpyxl.load_workbook(path, data_only=False)
    book = PurchaseBook(wb.active)

    with pytest.raises(DateColumnNotFoundError, match="2024-01-15"):
        book.require_date_column(dt.date(2024, 1, 15))


def test_purchase_book_missing_headers_names_the_table(tmp_path):
    path = tmp_path / "purchase.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["订单号", "型号"])  # 缺"采购日期"等必须的列
    wb.save(path)

    wb2 = openpyxl.load_workbook(path, data_only=False)
    with pytest.raises(HeaderNotFoundError, match="采购订单汇总表"):
        PurchaseBook(wb2.active)


# ---------------------------------------------------------------------------
# shipment_summary
# ---------------------------------------------------------------------------


def _write_summary_book(path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = [
        "采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
        "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号", "标签",
    ]
    ws.append(headers)
    ws.append(["PO-1", "M1", 5, 3, 15, None, "待定", "未发货", None, None, None, None, None, None, None, 7, "=+A2"])
    ws.append(["PO-2", "M2", 2, 3, 6, None, "待定", "未发货", None, None, None, None, None, None, None, 8, "=+A3"])
    ws.append([None] * 4 + ["=SUBTOTAL(9,E2:E3)"] + [None] * 12)  # 模拟表底的合计行
    wb.save(path)
    return wb


def test_shipment_summary_split_preserves_formatting(tmp_path):
    # 回归测试：真实文件是有格式的（字体、填充色、边框、行高），新插入的行要照抄模板行的格式，
    # 不能是一整行没有任何样式的空白格子；行高这种"整行"级别的设置也要跟着抄一份，不能丢。
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
               "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"]
    ws.append(headers)
    ws.append(["PO-1", "M1", 5, 3, 15, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    ws.append(["PO-2", "M2", 2, 3, 6, None, "待定", "未发货", None, None, None, None, None, None, None, None])

    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    thin = Border(top=Side(style="thin"), bottom=Side(style="thin"))
    for c in range(1, len(headers) + 1):
        ws.cell(row=2, column=c).fill = yellow
        ws.cell(row=2, column=c).border = thin
    ws.row_dimensions[2].height = 30
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2.active
    book = ShipmentSummaryBook(ws2)
    changes = book.apply_shipment("PO-1", "M1", 5, "ZD1", dt.date(2026, 9, 1))
    wb2.save(path)

    wb3 = openpyxl.load_workbook(path)
    ws3 = wb3.active
    new_row, pending_row = changes[0].new_row, changes[0].pending_row
    assert pending_row == 2  # 原来这一行原地转成已发货记录，不用挪位置
    assert new_row == 4  # 最后一行（PO-2）不是合计行，剩下的待定库存直接接在表格末尾
    assert ws3.cell(new_row, 1).fill.fgColor.rgb == "00FFFF00"
    assert ws3.cell(new_row, 1).border.top.style == "thin"
    assert ws3.row_dimensions[new_row].height == 30
    assert ws3.cell(pending_row, 1).fill.fgColor.rgb == "00FFFF00"
    assert ws3.cell(pending_row, 1).border.top.style == "thin"
    assert ws3.row_dimensions[pending_row].height == 30


def test_shipment_summary_reuses_trailing_blank_rows_instead_of_leaving_a_gap(tmp_path):
    # 回归测试：真实表格踩过的坑——表格末尾经常已经带着好几十行历史遗留的完全空白行（不是
    # 合计行，就是真的什么都没有）。之前的做法只看"最后一行"是不是空的，直接把新记录插在
    # 那个位置，完全没管上面那一大片同样空着、本该先被用掉的行——结果新记录跟用户原有的数据
    # 之间凭空隔出一段空行，看起来像"数据没接上"。现在应该优先把这些空白行填满，新记录跟着
    # 真实数据紧挨在一起。
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
               "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"]
    ws.append(headers)
    ws.append(["PO-1", "M1", 5, 3, 15, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    ws.append(["PO-2", "M2", 2, 3, 6, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    # 5 行历史遗留的完全空白行——真实表格里这种行往往还带着格式（边框/填充），只是没有值，
    # 这里用一下边框强迫 openpyxl 真的把这几行记进内部结构（不然 append 全 None 的行，
    # openpyxl 会当成什么都没发生，max_row 不会跟着变大，测不出真实场景）。
    thin = Border(top=Side(style="thin"))
    for r in range(3, 8):
        ws.cell(row=r + 1, column=1).border = thin
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2.active
    book = ShipmentSummaryBook(ws2)
    assert book._next_blank_row == 4  # 紧跟在真实数据（第 3 行）后面
    assert book._blank_row_limit == 9  # 5 行空白 -> 一直到第 8 行都能直接用

    changes = book.apply_shipment("PO-1", "M1", 5, "ZD1", dt.date(2026, 9, 1))
    new_row = changes[0].new_row
    assert new_row == 4  # 紧挨着真实数据插入，不是跳到第 8 行末尾之后
    assert ws2.cell(row=new_row, column=1).value == "PO-1"
    assert ws2.cell(row=new_row, column=5).value == 10  # 没扣完剩下的待定库存（15-5）

    # 再来一笔：应该接着用第 5 行（下一个空白行），不是又跳回第 4 行或者跳到最后
    changes2 = book.apply_shipment("PO-2", "M2", 2, "ZD1", dt.date(2026, 9, 1))
    assert changes2[0].new_row == 5
    assert ws2.cell(row=5, column=1).value == "PO-2"

    # 中间那些原本空白、还没被用到的行，要保持完全空白，不能被提前写进任何东西
    for r in (6, 7, 8):
        assert ws2.cell(row=r, column=1).value is None


def test_shipment_summary_split_reindexes_formulas(tmp_path):
    path = tmp_path / "summary.xlsx"
    _write_summary_book(path)
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    book = ShipmentSummaryBook(ws)

    changes = book.apply_shipment("PO-1", "M1", 5, "ZD1", dt.date(2026, 9, 1))
    assert len(changes) == 1
    change = changes[0]
    assert change.kind == "insert_new_pending"
    assert change.pending_row == 2  # 原来这一行原地转成已发货记录，不用挪位置
    assert change.new_row == 4  # 剩下的待定库存插在表格最下面（原来合计行所在的位置）
    assert change.pending_remaining_after == 10

    wb.save(path)
    wb2 = openpyxl.load_workbook(path, data_only=False)
    ws2 = wb2.active

    # 原来这一行原地转成已发货记录：数量/ZD/发货时间/状态是新值，标签公式还是指向自己这一行
    assert ws2.cell(row=2, column=5).value == 5  # 数量
    assert ws2.cell(row=2, column=6).value == "ZD1"
    assert ws2.cell(row=2, column=7).value == dt.datetime(2026, 9, 1)
    assert ws2.cell(row=2, column=17).value == "=+A2"

    # PO-2/M2 完全没被这次分摊碰到，还在原来的第 3 行，公式也没变
    assert ws2.cell(row=3, column=1).value == "PO-2"
    assert ws2.cell(row=3, column=17).value == "=+A3"

    # 新插入的待定行：数量是没扣完剩下的量，发货时间还是"待定"，标签公式指向自己这一行（第 4 行）
    assert ws2.cell(row=4, column=5).value == 10  # 数量
    assert ws2.cell(row=4, column=7).value == "待定"
    assert ws2.cell(row=4, column=17).value == "=+A4"

    # 合计行被顶到第 5 行，区间引用的结束边界从 E3 扩到 E4，把新插入的这一行也算进合计里；
    # 起始边界 E2 不变——起点那一行（PO-1）本来就没挪位置。
    assert ws2.cell(row=5, column=5).value == "=SUBTOTAL(9,E2:E4)"


def test_shipment_summary_writes_sku_into_label_column_not_huohao(tmp_path):
    # 回归测试：真实表里踩过的坑——"标签"这一列本该是运营发货计划表里原始的 SKU，但因为
    # 之前这一列没被当成要显式写的字段，是靠新行整行照抄模板行带过来的；而历史行的"标签"
    # 经常是"=+A<自己这行>"（这份 fixture 是镜像"采购单号"，真实表里更多是镜像"型号"/货号）
    # 这种公式，抄过去新行的"标签"也会跟着变成货号/别的东西，不是真正的 SKU。现在调用方传了
    # sku，就该显式覆盖成这个值，不管模板行原来是公式还是别的什么。
    path = tmp_path / "summary.xlsx"
    _write_summary_book(path)  # 第 17 列是"标签"，PO-1/PO-2 两行原来都是"=+A<row>"公式
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    book = ShipmentSummaryBook(ws)

    # 拆分：原来这一行（转成已发货记录）的标签要是传进去的 sku，不是模板行公式算出来的货号
    changes = book.apply_shipment("PO-1", "M1", 5, "ZD1", dt.date(2026, 9, 1), sku="WM-TD-348")
    pending_row = changes[0].pending_row
    new_row = changes[0].new_row
    assert ws.cell(row=pending_row, column=17).value == "WM-TD-348"
    # 新插入的待定行没被 _set_explicit_fields 碰过，标签公式原样不动（只是行号跟着挪到新行）
    assert ws.cell(row=new_row, column=17).value == "=+A4"

    # 原地转正（convert_in_place）：同一行的标签也要被覆盖成传进去的 sku
    changes2 = book.apply_shipment("PO-2", "M2", 6, "ZD1", dt.date(2026, 9, 1), sku="WM-TD-999")
    pending_row2 = changes2[0].pending_row
    assert ws.cell(row=pending_row2, column=17).value == "WM-TD-999"


def test_shipment_summary_convert_in_place_when_exact_match(tmp_path):
    path = tmp_path / "summary.xlsx"
    _write_summary_book(path)
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    book = ShipmentSummaryBook(ws)

    changes = book.apply_shipment("PO-2", "M2", 6, "ZD9", dt.date(2026, 9, 1))
    assert len(changes) == 1
    change = changes[0]
    assert change.kind == "convert_in_place"
    assert change.new_row is None
    assert ws.cell(row=3, column=5).value == 6
    assert ws.cell(row=3, column=7).value == dt.datetime(2026, 9, 1)
    assert ws.cell(row=3, column=8).value == "未发货"


def test_shipment_summary_ignores_pending_rows_with_non_shipping_status(tmp_path):
    # 回归测试：真实表里有极少数行"发货时间"还留着"待定"，但"状态"已经被人工改成"已取消"
    # 或"无库存"（改状态的时候忘了把发货时间一起改掉）。这种行不是真的能扣的库存，只看
    # "发货时间=待定"会把它们错当成待定库存去扣，必须同时要求"状态=未发货"才算数。
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
               "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"]
    ws.append(headers)
    ws.append(["PO-1", "M1", 1, 3, 3, None, "待定", "已取消", None, None, None, None, None, None, None, None])
    ws.append(["PO-1", "M1", 5, 3, 15, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2.active
    book = ShipmentSummaryBook(ws2)

    # 只有第 3 行（状态=未发货）才是真正的待定行；第 2 行（已取消）不该出现在候选里
    assert book.pending_rows("PO-1", "M1") == [3]
    assert book.total_pending_quantity("PO-1", "M1") == 15

    changes = book.apply_shipment("PO-1", "M1", 15, "ZD1", dt.date(2026, 9, 1))
    assert len(changes) == 1
    assert changes[0].kind == "convert_in_place"
    assert changes[0].pending_row == 3
    # 已取消的那一行必须原封不动，不能被写入任何发货信息
    assert ws2.cell(row=2, column=7).value == "待定"
    assert ws2.cell(row=2, column=8).value == "已取消"
    assert ws2.cell(row=2, column=5).value == 3


def test_shipment_summary_blank_fields_actually_clear_stale_values(tmp_path):
    # 回归测试：真实数据里"待定"行经常已经带着"编号"/"备注"这些字段的历史值（比如"无库存9"）——
    # _blank_fields 之前用 ws.cell(row, col, value=None) 清空，这个写法在 openpyxl 里是个坑：
    # value 只有不是 None 才会真的赋值，传 None 等于没传，格子会原样留着旧值。这里让待定行带上
    # 非空的 编号/备注，转正/拆分之后必须变成空，不能把旧备注误当成新记录的状态。
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
               "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"]
    ws.append(headers)
    ws.append(["PO-1", "M1", 5, 3, 15, None, "待定", "未发货", "旧仓库", "旧FBA", "旧追踪", "无库存9", "旧货代", "旧出货单", None, 3])
    ws.append(["PO-2", "M2", 2, 3, 6, None, "待定", "未发货", "旧仓库2", None, None, "另一条旧备注", None, None, None, 8])
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2.active
    book = ShipmentSummaryBook(ws2)

    # 拆分：新插入的待定行是从原来那行复制出来的，旧的 仓库/FBA/追踪/备注/货代/出货单/编号
    # 都不该带过去（还没决定怎么发，不该继承上一轮遗留的旧备注）
    changes = book.apply_shipment("PO-1", "M1", 5, "ZD1", dt.date(2026, 9, 1))
    new_row = changes[0].new_row
    for col in (9, 10, 11, 12, 13, 14, 16):  # 仓库/FBA ID/追踪编号/备注/货代/出货单号/编号
        assert ws2.cell(row=new_row, column=col).value is None, f"col {col} 应该清空"

    # convert_in_place：原地转正的那一行自己带的旧值也要被清掉
    changes2 = book.apply_shipment("PO-2", "M2", 6, "ZD2", dt.date(2026, 9, 1))
    pending_row = changes2[0].pending_row
    for col in (9, 10, 11, 12, 13, 14, 16):
        assert ws2.cell(row=pending_row, column=col).value is None, f"col {col} 应该清空"


def test_shipment_summary_missing_box_capacity_still_updates_boxes(tmp_path):
    # 回归测试：_set_explicit_fields/拆分剩余量那两处算出来的箱数在箱容缺失时会是 None——
    # 之前同样因为 ws.cell(..., value=None) 是 no-op，箱容缺失的行拆分/转正之后箱数格子会
    # 留着旧值，跟新的「数量」对不上。这里箱容留空，箱数原来的旧值应该被清掉，不能留着旧数字。
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
               "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"])
    ws.append(["PO-1", "M1", 999, None, 15, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2.active
    book = ShipmentSummaryBook(ws2)

    changes = book.apply_shipment("PO-1", "M1", 5, "ZD1", dt.date(2026, 9, 1))
    new_row, pending_row = changes[0].new_row, changes[0].pending_row
    # 箱容缺失，箱数算不出来，应该是 None，不能留着模板行的旧箱数 999
    assert ws2.cell(row=new_row, column=3).value is None
    assert ws2.cell(row=pending_row, column=3).value is None
    # pending_row 原地转成已发货记录（数量=这次发的 5），new_row 是没扣完剩下的待定（数量=10）
    assert ws2.cell(row=pending_row, column=5).value == 5
    assert ws2.cell(row=new_row, column=5).value == 10


def test_shipment_summary_pending_row_not_found_raises(tmp_path):
    path = tmp_path / "summary.xlsx"
    _write_summary_book(path)
    wb = openpyxl.load_workbook(path)
    book = ShipmentSummaryBook(wb.active)
    with pytest.raises(Exception):
        book.apply_shipment("PO-不存在", "M1", 1, "ZD", dt.date(2026, 9, 1))


def test_shipment_summary_skips_zero_quantity_sibling_row(tmp_path):
    # 回归测试：真实数据里出现过同一个采购单号+型号同时有两行待定——一行剩 36 个、一行是
    # 历史遗留的 0 个（比如对应的采购记录后来被清零了）。之前的实现只按采购单号+型号找，
    # 谁先出现在表里就抢到谁，曾经把 21 个写进了那条本该是 0 的待定行，还把它的箱数从 0
    # 改掉了；真正有货的那条反而没动。现在应该自动跳过 0 数量的行，从有货的那行扣。
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
               "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"])
    # 先放那条 0 数量的（模拟它在表里排在前面，最容易被"谁先找到用谁"的旧逻辑误伤）
    ws.append(["PO-DUP", "M1", 0, 3, 0, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    ws.append(["PO-DUP", "M1", 12, 3, 36, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    book = ShipmentSummaryBook(wb2.active)

    changes = book.apply_shipment("PO-DUP", "M1", 21, "ZD1", dt.date(2026, 9, 1))
    assert len(changes) == 1
    assert changes[0].kind == "insert_new_pending"
    assert changes[0].pending_remaining_after == 15

    ws2 = wb2.active
    # 那条 0 数量的待定行必须原封不动，一个字段都不能被碰
    zero_row_values = [ws2.cell(row=2, column=c).value for c in range(1, 9)]
    assert zero_row_values == ["PO-DUP", "M1", 0, 3, 0, None, "待定", "未发货"]


def test_shipment_summary_drains_across_multiple_pending_rows(tmp_path):
    # 同一个采购单号+型号同时有好几行待定是正常状态（都是同一批还没决定去哪的库存），要发的
    # 数量比其中一行多的话，应该继续从下一行扣，不是报错——两行加起来（10+20=30）够 25 就行。
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
               "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"])
    ws.append(["PO-MULTI", "M1", 10, 1, 10, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    ws.append(["PO-MULTI", "M1", 20, 1, 20, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    book = ShipmentSummaryBook(wb2.active)

    changes = book.apply_shipment("PO-MULTI", "M1", 25, "ZD1", dt.date(2026, 9, 1))
    assert len(changes) == 2
    assert changes[0].kind == "convert_in_place"  # 第一行 10 个正好扣完
    assert changes[0].quantity == 10
    assert changes[1].kind == "insert_new_pending"  # 第二行扣 15，剩 5 还待定
    assert changes[1].quantity == 15
    assert changes[1].pending_remaining_after == 5

    assert book.total_pending_quantity("PO-MULTI", "M1") == 5


def test_shipment_summary_total_exceeding_all_pending_rows_raises(tmp_path):
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
               "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号"])
    ws.append(["PO-DUP", "M1", 0, 3, 0, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    ws.append(["PO-DUP", "M1", 12, 3, 36, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    wb.save(path)
    wb2 = openpyxl.load_workbook(path)
    book = ShipmentSummaryBook(wb2.active)

    # 两行加起来一共 36，要发 999 应该直接报错，不能瞎猜
    with pytest.raises(Exception):
        book.apply_shipment("PO-DUP", "M1", 999, "ZD1", dt.date(2026, 9, 1))


def test_shipment_summary_quantity_exceeding_pending_raises(tmp_path):
    path = tmp_path / "summary.xlsx"
    _write_summary_book(path)
    wb = openpyxl.load_workbook(path)
    book = ShipmentSummaryBook(wb.active)
    # PO-1/M1 待定数量是 15（见 _write_summary_book），要发 999 个应该直接报错，而不是被当成
    # "刚好发完"原地转正、把数字写错
    with pytest.raises(Exception):
        book.apply_shipment("PO-1", "M1", 999, "ZD1", dt.date(2026, 9, 1))


def test_shipment_summary_sync_auto_filter_extends_stale_range(tmp_path):
    # 回归测试：真实表里踩过的坑——AutoFilter 的范围是写死在文件里的固定区间，插入新行/
    # 转正已有行都不会让它自动跟着扩大。写完之后不补这一步的话，数据其实是对的，但在 Excel
    # 里拿筛选框去找新写的记录会找不到（超出筛选范围看不见），容易被误以为没写进去。这里
    # 特意让初始筛选范围比数据本身还窄（模拟真实表里筛选范围早就过期的情况），确认调用
    # sync_auto_filter() 之后范围会扩大到覆盖所有数据，但不会缩小已经比数据边界更大的范围。
    path = tmp_path / "summary.xlsx"
    _write_summary_book(path)  # 4 行数据（含表底合计行），17 列
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    ws.auto_filter.ref = "A1:C2"  # 故意设得比实际数据范围窄很多
    book = ShipmentSummaryBook(ws)

    book.apply_shipment("PO-1", "M1", 5, "ZD1", dt.date(2026, 9, 1))  # 会在表格最下面插一行
    book.sync_auto_filter()

    from openpyxl.utils.cell import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(ws.auto_filter.ref)
    assert min_row == 1  # 只扩大，原来的起点（第 1 行）不会被抬高
    assert max_row == book._max_row  # 覆盖到新插入行之后的最后一行
    assert max_col >= book._max_column

    # 范围已经比实际数据边界更大（比如人工手动扩过）的话，不应该被这一步缩小
    ws.auto_filter.ref = f"A1:Z{book._max_row + 50}"
    book.sync_auto_filter()
    assert ws.auto_filter.ref == f"A1:Z{book._max_row + 50}"


def test_shipment_summary_missing_headers_names_the_table(tmp_path):
    path = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["采购单号", "型号"])  # 缺"发货时间"等必须的列
    wb.save(path)

    wb2 = openpyxl.load_workbook(path)
    with pytest.raises(HeaderNotFoundError, match="发货计划汇总表"):
        ShipmentSummaryBook(wb2.active)


# ---------------------------------------------------------------------------
# planner + diff：串起来的小型集成测试（纯合成数据，不依赖真实文件）
# ---------------------------------------------------------------------------


def test_planner_and_diff_end_to_end(tmp_path):
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    # 发货计划汇总表里的待定行要跟采购汇总表对得上号（同一个采购单号+型号），这里手写一份，
    # 不能直接借用 _write_summary_book 那份（用的是 PO-1/PO-2，跟这里的 PO-EARLY 对不上）
    summary_path = tmp_path / "summary.xlsx"
    summary_setup_wb = openpyxl.Workbook()
    summary_ws = summary_setup_wb.active
    headers = [
        "采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
        "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号",
    ]
    summary_ws.append(headers)
    summary_ws.append(["PO-EARLY", "M1", 27, 3, 80, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    summary_setup_wb.save(summary_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    summary_book = ShipmentSummaryBook(summary_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=30, destination_label="ZD1", source_row=2)
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors

    result = run_and_capture_diff(plan, purchase_book, summary_book)
    assert result.purchase.before_rows[0]["未出货数量"] == 80
    assert result.purchase.after_rows[0]["未出货数量"] == 50
    assert len(result.summary.before_rows) == 1
    assert len(result.summary.after_rows) == 2  # 部分发货：新行 + 剩余待定行


def test_diff_preview_does_not_duplicate_after_rows_across_sibling_pending_rows(tmp_path):
    # 回归测试：同一个采购单号+型号同时有两行"待定"，一笔出货正好吃掉第一行全部、又部分
    # 吃掉第二行——预览应该是"第一行删除 + 它对应的 1 条新增"、"第二行删除 + 它对应的 2 条
    # 新增"，而不是把这次操作产生的全部 3 条新增行重复挂在每一个被删除的原始行下面。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    summary_path = tmp_path / "summary.xlsx"
    summary_setup_wb = openpyxl.Workbook()
    summary_ws = summary_setup_wb.active
    headers = [
        "采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
        "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号",
    ]
    summary_ws.append(headers)
    # 两行待定：一行 10 个（会被整行转正），一行 20 个（会被扣 15，剩 5 还待定）
    summary_ws.append(["PO-EARLY", "M1", 10, 1, 10, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    summary_ws.append(["PO-EARLY", "M1", 20, 1, 20, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    summary_setup_wb.save(summary_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    summary_book = ShipmentSummaryBook(summary_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=25, destination_label="ZD1", source_row=2)
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors

    result = run_and_capture_diff(plan, purchase_book, summary_book)

    assert len(result.summary.before_rows) == 2
    assert len(result.summary.after_rows) == 3  # 1 条整行转正 + (1 条新发货行 + 1 条剩余待定行)

    before_groups = [row[GROUP_KEY] for row in result.summary.before_rows]
    assert len(set(before_groups)) == 2  # 两条 before 行分属两个不同的组，不能被合并

    after_by_group: dict = {}
    for row in result.summary.after_rows:
        after_by_group.setdefault(row[GROUP_KEY], []).append(row)

    # 第一行（原 10 个，整行转正）只对应 1 条新增
    group_of_10 = next(row[GROUP_KEY] for row in result.summary.before_rows if row["数量"] == 10)
    assert len(after_by_group[group_of_10]) == 1

    # 第二行（原 20 个，部分扣减）对应 2 条新增：新发货的 15 + 剩余待定的 5
    group_of_20 = next(row[GROUP_KEY] for row in result.summary.before_rows if row["数量"] == 20)
    assert len(after_by_group[group_of_20]) == 2
    assert {row["数量"] for row in after_by_group[group_of_20]} == {15, 5}


def test_diff_preview_group_tracks_leftover_row_consumed_by_a_later_change(tmp_path):
    # 回归测试：同一分组的"剩余待定行"如果之后又被另一个 item 继续扣（这一批发货计划里有
    # 两条记录都指向同一个采购单号+型号），预览不应该还显示一条早就不存在了的"剩余待定行"，
    # 应该显示两条真正落地的发货记录。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    summary_path = tmp_path / "summary.xlsx"
    summary_setup_wb = openpyxl.Workbook()
    summary_ws = summary_setup_wb.active
    headers = [
        "采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
        "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号",
    ]
    summary_ws.append(headers)
    summary_ws.append(["PO-EARLY", "M1", 30, 1, 30, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    summary_setup_wb.save(summary_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    summary_book = ShipmentSummaryBook(summary_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    # 两条计划行都指向同一个采购单号+型号：第一条扣 10（剩 20 还待定），第二条再扣 20（正好扣完）
    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=10, destination_label="ZD1", source_row=2),
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=20, destination_label="ZD1", source_row=3),
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors

    result = run_and_capture_diff(plan, purchase_book, summary_book)

    assert len(result.summary.before_rows) == 1
    # 应该是 2 条发货记录（10 个 + 20 个），不应该还多一条"剩余 20 待定"（那一行早就被第二条
    # change 继续扣完了，不再存在）
    assert len(result.summary.after_rows) == 2
    assert {row["数量"] for row in result.summary.after_rows} == {10, 20}
    assert all(row[GROUP_KEY] == result.summary.before_rows[0][GROUP_KEY] for row in result.summary.after_rows)


def test_planner_skips_shortfall_item_without_blocking_batch(tmp_path):
    # 余量不够不再阻塞整批——只把这一行跳过（记进 skip_reason，不写任何东西），同一批里其它
    # 行照常处理。跳过的时候要把这次分摊已经占用的余量退回去，不然会白白占掉后面同货号其它
    # 行本来还能分到的库存（这里第二条行请求的 50 个，如果没退回去会跟着一起分不到）。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=9999, destination_label="ZD1", source_row=2),
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=50, destination_label="ZD1", source_row=3),
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors

    assert len(plan.skipped_items) == 1
    skipped = plan.skipped_items[0]
    assert skipped.line.quantity == 9999
    assert skipped.allocations == []
    assert "还差" in skipped.skip_reason

    processed = next(item for item in plan.items if item.line.quantity == 50)
    assert processed.skip_reason is None
    assert sum(a.quantity for a in processed.allocations) == 50  # 余量被正确退回，第二条行分到满额

    apply_plan_purchase_only(plan, purchase_book)  # 不再抛错
    date_col = purchase_book.require_date_column(dt.date(2026, 9, 1))
    early_row = next(r for r in purchase_book.rows if r.order_no == "PO-EARLY")
    assert purchase_wb.active.cell(row=early_row.row_index, column=date_col).value == 50


def test_planner_prioritizes_amazon_then_walmart_then_overseas_on_shortfall(tmp_path):
    # 库存不够分给所有平台的时候，缺口应该落在海外仓身上——不管这几行在输入列表里的先后
    # 顺序，build_plan 都要按亚马逊>沃尔玛>海外仓重新排过再分摊。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([
        "订单号", "采购日期", "型号", "订单数量", "数量单位",
        dt.datetime(2026, 9, 1), "未出货数量",
    ])
    ws.append([None, None, None, None, None, "出货时间", None])
    ws.append(["PO-1", dt.datetime(2026, 1, 1), "M1", 30, "pcs", None, "=D3-SUM(F3:F3)"])
    wb.save(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    # 故意把输入顺序反着排（海外仓在前、亚马逊在后），确认真正生效的是模板优先级，不是
    # 输入列表原来的先后顺序。
    lines = [
        PlanLine(
            zd="LO-RK", sku_kind="RK", sku="RK-1", quantity=30, destination_label="LO-RK",
            source_row=2, template_type="overseas",
        ),
        PlanLine(
            zd="US", sku_kind="AMZ", sku="AMZ-1", quantity=30, destination_label="US",
            source_row=2, template_type="amazon",
        ),
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors

    amazon_item = next(i for i in plan.items if i.line.template_type == "amazon")
    overseas_item = next(i for i in plan.items if i.line.template_type == "overseas")
    assert amazon_item.skip_reason is None
    assert sum(a.quantity for a in amazon_item.allocations) == 30
    assert overseas_item.skip_reason is not None
    assert overseas_item.allocations == []


def test_write_skipped_items_report_creates_file_next_to_first_plan(tmp_path):
    from modules.logistics.shipment_plan_apply.planner import PlanItem, write_skipped_items_report
    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    plan_dir = tmp_path / "plans"
    plan_dir.mkdir()
    first_plan_path = plan_dir / "亚马逊计划.xlsx"

    # 没有任何跳过的行就不该生成文件
    assert write_skipped_items_report([], first_plan_path) is None

    skipped_line = PlanLine(
        zd="LO-RK", sku_kind="RK", sku="RK-9", quantity=99, destination_label="LO-RK",
        source_row=5, source_file="海外仓计划.xlsx", template_type="overseas",
    )
    skipped_item = PlanItem(line=skipped_line, huohao="TD-9", skip_reason="货号「TD-9」还差 10 个，已跳过")

    out_path = write_skipped_items_report([skipped_item], first_plan_path)
    assert out_path is not None
    assert out_path.parent == plan_dir
    assert out_path.name.startswith("亚马逊计划-异常数据-")

    wb = openpyxl.load_workbook(out_path)
    ws = wb.active
    header = [c.value for c in ws[1]]
    assert header == ["来源文件", "表格行数", "平台", "SKU 种类", "SKU", "货号", "目的地/ZD", "需要数量", "说明"]
    row = [c.value for c in ws[2]]
    assert row == ["海外仓计划.xlsx", 5, "海外仓", "RK", "RK-9", "TD-9", "LO-RK", 99, "货号「TD-9」还差 10 个，已跳过"]


def test_build_plan_reports_missing_date_column_instead_of_inserting(tmp_path):
    # 日期列不存在的时候，build_plan 要在这里就报出来（进不了 apply 阶段），不能像早期版本
    # 那样在 apply 的时候现场插一列——见 purchase_book.py 顶部说明。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path, include_2026_09_01_col=False)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)
    purchase_max_column_before = purchase_wb.active.max_column

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=30, destination_label="ZD1", source_row=2)
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))

    assert plan.has_blocking_errors
    assert plan.items == []
    assert any("手动加好这一天的日期列" in e for e in plan.parse_errors)
    # 没有任何东西被写进去，列数也没变——报错发生在任何写操作之前。
    assert purchase_wb.active.max_column == purchase_max_column_before


def test_apply_plan_purchase_only_accumulates_same_sku_different_zd(tmp_path):
    # "采购订单分摊更新"（只写采购订单汇总表，不碰发货计划汇总表）：同一个货号在这一批里
    # 可能因为分属不同的目的地（不同 ZD）拆成好几条 PlanItem，采购订单汇总表的日期列本来
    # 就是"这一天一共发了多少"的汇总，不分 ZD——两笔分摊写到同一个日期列，应该叠加在一起
    # （30+20=50），不能后一笔把前一笔覆盖掉。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    lines = [
        PlanLine(zd="CK-WM", sku_kind="RK", sku="RK-1", quantity=30, destination_label="US", source_row=2),
        PlanLine(zd="LO-WM", sku_kind="RK", sku="RK-1", quantity=20, destination_label="CA", source_row=3),
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors
    assert plan.total_allocations == 2  # 两条 PlanItem，各自分摊一次

    apply_plan_purchase_only(plan, purchase_book, None)

    date_col = purchase_book.require_date_column(dt.date(2026, 9, 1))
    early_row = next(r for r in purchase_book.rows if r.order_no == "PO-EARLY")
    # 两笔分摊（先扣早的订单）都落在同一行、同一个日期列，写入的量是 30+20=50，不是后一笔
    # 把前一笔覆盖成 20。
    assert purchase_wb.active.cell(row=early_row.row_index, column=date_col).value == 50


def test_run_and_capture_diff_purchase_only_reports_only_purchase_changes(tmp_path):
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=30, destination_label="ZD1", source_row=2)
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors

    diff_table = run_and_capture_diff_purchase_only(plan, purchase_book)
    assert diff_table.before_rows[0]["未出货数量"] == 80
    assert diff_table.after_rows[0]["未出货数量"] == 50


def test_apply_plan_summary_only_does_not_touch_purchase_book(tmp_path):
    # "发货计划汇总表更新"：purchase_book 只用来算分摊（决定这个货号该扣哪几笔采购订单），
    # 不该往它里面写任何东西——不插日期列、不写累计出货量，调用方也不需要保存/备份它。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)
    purchase_max_column_before = purchase_wb.active.max_column

    summary_path = tmp_path / "summary.xlsx"
    summary_setup_wb = openpyxl.Workbook()
    summary_ws = summary_setup_wb.active
    headers = [
        "采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
        "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号",
    ]
    summary_ws.append(headers)
    summary_ws.append(["PO-EARLY", "M1", 27, 3, 80, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    summary_setup_wb.save(summary_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    summary_book = ShipmentSummaryBook(summary_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=30, destination_label="ZD1", source_row=2)
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors

    apply_plan_summary_only(plan, summary_book)

    # 采购订单汇总表完全没被改动：列数没变，2026-09-01 这一列（fixture 里本来就有）也
    # 还是空的——apply_plan_summary_only 不写累计出货量，只是拿 purchase_book 算分摊。
    assert purchase_wb.active.max_column == purchase_max_column_before
    date_col = purchase_book.find_date_column(dt.date(2026, 9, 1))
    assert all(
        purchase_wb.active.cell(row=r.row_index, column=date_col).value is None for r in purchase_book.rows
    )

    # 发货计划汇总表这边确实被扣了
    assert summary_book.total_pending_quantity("PO-EARLY", "M1") == 50


def test_run_and_capture_diff_summary_only_reports_only_summary_changes(tmp_path):
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    summary_path = tmp_path / "summary.xlsx"
    summary_setup_wb = openpyxl.Workbook()
    summary_ws = summary_setup_wb.active
    headers = [
        "采购单号", "型号", "箱数", "箱容", "数量", "ZD", "发货时间", "状态",
        "仓库", "FBA ID", "追踪编号", "备注", "货代", "出货单号", "so", "编号",
    ]
    summary_ws.append(headers)
    summary_ws.append(["PO-EARLY", "M1", 27, 3, 80, None, "待定", "未发货", None, None, None, None, None, None, None, None])
    summary_setup_wb.save(summary_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    summary_book = ShipmentSummaryBook(summary_wb.active)

    from modules.logistics.shipment_plan_apply.shipment_templates import PlanLine

    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=30, destination_label="ZD1", source_row=2)
    ]
    plan = build_plan(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))
    assert not plan.has_blocking_errors

    diff_table = run_and_capture_diff_summary_only(plan, summary_book)
    assert len(diff_table.before_rows) == 1
    assert diff_table.before_rows[0]["数量"] == 80
    assert len(diff_table.after_rows) == 2  # 部分发货：新行 + 剩余待定行
    assert {row["数量"] for row in diff_table.after_rows} == {30, 50}


def test_build_plan_from_recorded_allocations_replays_already_written_purchase_columns(tmp_path):
    # 模拟真实场景：先用「采购订单分摊更新」把这一批写进采购订单汇总表并存盘，「发货计划
    # 汇总表更新」拿到的是这份已经写过的采购表——这时候不该再按余量现算一遍分摊（余量已经
    # 被这一批扣过，重新算会报"余量不够"，但其实是这一批本来就该分摊到的量）。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path)

    # 第一步：模拟「采购订单分摊更新」已经跑过——两条不同 ZD 的发货计划行（60 + 30），
    # 按余量从早到晚分摊（PO-EARLY 剩 80，PO-LATE 剩 50），写入 2026-9-1 这一列，存盘。
    setup_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    setup_book = PurchaseBook(setup_wb.active)
    date_col = setup_book.require_date_column(dt.date(2026, 9, 1))
    for qty in (60, 30):
        outcome = setup_book.allocate("M1", qty)
        assert outcome.shortfall == 0
        for allocation in outcome.allocations:
            setup_book.write_allocation(allocation, date_col)
    setup_wb.save(purchase_path)
    # PO-EARLY 这一天被写了 80（60 之后剩 20 给第二笔用满），PO-LATE 被写了 10。
    early_row = next(r for r in setup_book.rows if r.order_no == "PO-EARLY")
    late_row = next(r for r in setup_book.rows if r.order_no == "PO-LATE")
    assert setup_wb.active.cell(row=early_row.row_index, column=date_col).value == 80
    assert setup_wb.active.cell(row=late_row.row_index, column=date_col).value == 10

    # 第二步：全新加载这份"已经用过"的采购表（模拟另一次工具运行、全新的 PurchaseBook 实例），
    # 拿它去重放「发货计划汇总表更新」该匹配哪几笔采购订单——不能再用 allocate()，因为余量
    # 早就被上面这一批用掉了，会误报"余量不够"。
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)
    assert purchase_book.by_model["M1"][0].remaining == 0  # PO-EARLY 余量确实已经是 0

    lines = [
        PlanLine(zd="CK-WM", sku_kind="RK", sku="RK-1", quantity=60, destination_label="US", source_row=2),
        PlanLine(zd="LO-WM", sku_kind="RK", sku="RK-1", quantity=30, destination_label="CA", source_row=3),
    ]
    plan = build_plan_from_recorded_allocations(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))

    assert not plan.has_blocking_errors, [e for i in plan.items for e in i.errors] + plan.parse_errors
    assert plan.total_allocations == 3  # 60 拆成 PO-EARLY 60；30 拆成 PO-EARLY 20 + PO-LATE 10
    first_allocations = [(a.row.order_no, a.quantity) for a in plan.items[0].allocations]
    second_allocations = [(a.row.order_no, a.quantity) for a in plan.items[1].allocations]
    assert first_allocations == [("PO-EARLY", 60)]
    assert second_allocations == [("PO-EARLY", 20), ("PO-LATE", 10)]


def test_build_plan_from_recorded_allocations_errors_when_date_column_missing(tmp_path):
    # 采购表还没被「采购订单分摊更新」用过（没有这一天的日期列）——不该去按余量现算，
    # 应该明确提示"先用「采购订单分摊更新」把这一批写进去"。
    product_path = tmp_path / "product.xlsx"
    _write_product_info(product_path, [("AMZ-1", "RK-1", "M1")])
    lookup = load_product_lookup(product_path)

    purchase_path = tmp_path / "purchase.xlsx"
    _write_purchase_book(purchase_path, include_2026_09_01_col=False)
    purchase_wb = openpyxl.load_workbook(purchase_path, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    lines = [
        PlanLine(zd="ZD1", sku_kind="RK", sku="RK-1", quantity=30, destination_label="ZD1", source_row=2)
    ]
    plan = build_plan_from_recorded_allocations(lines, [], lookup, purchase_book, dt.date(2026, 9, 1))

    assert plan.has_blocking_errors
    assert plan.items == []
    assert any("先用「采购订单分摊更新」" in e for e in plan.parse_errors)


# ---------------------------------------------------------------------------
# 真实数据集成测试（跑起来比较慢，日常开发用 -m "not slow" 跳过）
# ---------------------------------------------------------------------------


@pytest.mark.skipif(not REAL_DATA_DIR.exists(), reason="需要本机真实的物流仓库数据文件")
def test_real_walmart_plan_allocates_against_real_purchase_book(tmp_path):
    import shutil

    lookup = load_product_lookup(REAL_DATA_DIR / "在售产品信息总表(测试用).xlsx")

    plan_path = REAL_DATA_DIR / "CK-Walmart9.1发货计划.xlsx"
    parsed = parse_shipment_plan(plan_path, "9.1")
    assert parsed.template_type == "walmart"
    assert not parsed.errors

    purchase_tmp = tmp_path / "purchase.xlsx"
    shutil.copy(REAL_DATA_DIR / "采购订单汇总表(测试用).xlsx", purchase_tmp)
    purchase_wb = openpyxl.load_workbook(purchase_tmp, data_only=False)
    purchase_book = PurchaseBook(purchase_wb.active)

    plan = build_plan(parsed.lines, parsed.errors, lookup, purchase_book, dt.date(2026, 9, 1))
    # 这份真实数据里 TD-640 这条已知缺货，整批应该被挡住
    assert plan.has_blocking_errors
    assert any("TD-RZ-585" in e for item in plan.items for e in item.errors)
