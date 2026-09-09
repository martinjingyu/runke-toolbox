import datetime as dt
from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import PatternFill
from openpyxl.worksheet.formula import ArrayFormula
from PySide6.QtCore import QSettings

from modules.logistics.purchase_order_import.order_file import OrderFileError, parse_order_file
from modules.logistics.purchase_order_import.planner import build_plan, apply_plan
from modules.logistics.purchase_order_import.supplier_codes import SupplierCodeStore


# ---------------------------------------------------------------------------
# order_file.parse_order_file
# ---------------------------------------------------------------------------


def _write_order_file(path: Path, order_no="SX-2609209", supplier="东莞市盛鑫灯饰有限公司", rows=None) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "合同 (2)"
    ws.append(["香港闰科有限公司"])
    ws.append(["采购订单"])
    ws.append([None, None, None, None, f"订单编号：{order_no}   PO：PO#{order_no}"])
    ws.append([
        "采购商（甲方）：香港闰科有限公司", None, None, None, "地址：xxx", "联系方式：", "139", None,
        "采购日期：", dt.datetime(2026, 9, 2),
    ])
    ws.append([f"供应商（乙方）：{supplier}", None, None, None, "地址：xxx", "联系方式：", None, None, "交货日期：", "详见附件"])
    ws.append(["序号", "sku", "产品名称", "图片IMG", "描述", "包装方式及尺寸", "数量", "单价", "交货日期", "备注"])
    for i, (sku, name, qty, delivery) in enumerate(rows or [], start=1):
        ws.append([i, sku, name, None, None, None, qty, 10, delivery, None])
    wb.save(path)


def test_parse_order_file_extracts_header_fields_and_lines(tmp_path):
    path = tmp_path / "order.xlsx"
    _write_order_file(
        path,
        rows=[
            ("TD-CY-410", "雅筑黑色五金台灯", 150, dt.datetime(2026, 9, 29)),
            ("TD-LO-92", "白门框台灯2p", 180, dt.datetime(2026, 10, 7)),
        ],
    )
    order = parse_order_file(path)
    assert order.order_no == "SX-2609209"
    assert order.supplier_name == "东莞市盛鑫灯饰有限公司"
    assert order.purchase_date == dt.date(2026, 9, 2)
    assert not order.errors
    assert len(order.lines) == 2
    assert order.lines[0].model == "TD-CY-410"
    assert order.lines[0].quantity == 150
    assert order.lines[0].delivery_date == dt.date(2026, 9, 29)


def test_parse_order_file_skips_blank_sku_rows(tmp_path):
    path = tmp_path / "order.xlsx"
    _write_order_file(path, rows=[("TD-CY-410", "灯", 10, dt.datetime(2026, 9, 29))])
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    ws.append([3, None, None, None, None, None, None, None, None, None])  # sku 空，应该跳过
    wb.save(path)

    order = parse_order_file(path)
    assert len(order.lines) == 1


def test_parse_order_file_missing_order_no_raises(tmp_path):
    path = tmp_path / "order.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["没有订单编号的表格"])
    ws.append(["序号", "sku", "产品名称", "图片IMG", "描述", "包装方式及尺寸", "数量", "单价", "交货日期", "备注"])
    ws.append([1, "TD-1", "灯", None, None, None, 10, 5, dt.datetime(2026, 1, 1), None])
    wb.save(path)

    with pytest.raises(OrderFileError):
        parse_order_file(path)


def test_parse_order_file_missing_supplier_leaves_none_and_notes(tmp_path):
    path = tmp_path / "order.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([None, None, None, None, "订单编号：AB-001"])
    ws.append(["序号", "sku", "产品名称", "图片IMG", "描述", "包装方式及尺寸", "数量", "单价", "交货日期", "备注"])
    ws.append([1, "TD-1", "灯", None, None, None, 10, 5, dt.datetime(2026, 1, 1), None])
    wb.save(path)

    order = parse_order_file(path)
    assert order.order_no == "AB-001"
    assert order.supplier_name is None
    assert order.purchase_date is None
    assert any("供应商" in e for e in order.errors)
    assert any("采购日期" in e for e in order.errors)


# ---------------------------------------------------------------------------
# supplier_codes.SupplierCodeStore
# ---------------------------------------------------------------------------


@pytest.fixture()
def settings(tmp_path):
    path = str(tmp_path / "settings.ini")
    return QSettings(path, QSettings.Format.IniFormat)


def test_supplier_code_store_roundtrip(settings):
    store = SupplierCodeStore(settings)
    assert store.mapping() == {}
    assert store.resolve("东莞市盛鑫灯饰有限公司") is None

    store.set_mapping({"东莞市盛鑫灯饰有限公司": "SX", "GH工厂": "GH"})
    assert store.mapping() == {"东莞市盛鑫灯饰有限公司": "SX", "GH工厂": "GH"}
    assert store.resolve("东莞市盛鑫灯饰有限公司") == "SX"
    assert store.resolve(None) is None

    store.set_mapping({})
    assert store.mapping() == {}


# ---------------------------------------------------------------------------
# planner.build_plan / apply_plan
# ---------------------------------------------------------------------------


def _write_purchase_summary(path: Path) -> None:
    # PurchaseBook 要求表头里有「数量单位」和「未出货数量」，且两者之间至少留一列日期列
    # （见 purchase_book.py），所以测试表里也要把这几列凑齐，不能只写 build_plan 用得到的
    # 那几列；PurchaseBook 还假设表头下一行是"出货时间"这种子表头行，真正的数据从表头
    # 再往下第二行才开始（见 purchase_book.py 的 sub_header_row / _load_rows），测试表结构
    # 要跟真实表一致（第1行标题、第2行表头、第3行子表头、第4行起才是数据），不然
    # PurchaseBook 会读不到任何一行数据。
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "采购订单汇总"
    ws.append([None] * 12)  # 第1行：标题/合计行占位
    ws.append([
        "序号", "订单号", "采购日期", "交货日期", "供应商名称", "店铺", "型号", "产品名称",
        "订单数量", "数量单位", "出货日期占位", "未出货数量",
    ])
    ws.append([None] * 10 + ["出货时间", None])  # 第3行：子表头
    ws.append([
        "001", "GH-2501009", dt.datetime(2025, 1, 7), dt.datetime(2025, 3, 17), "GH", None,
        "TD-RZ-419", "简约花瓶灰色树脂台灯", 180, "pcs", None, "=I4",
    ])
    # 给型号格子上个底色，模拟真实表格逐行有格式——新增行应该跟着抄这个格式
    ws.cell(row=4, column=7).fill = PatternFill("solid", fgColor="FFFF00")
    wb.save(path)


def _write_shipment_summary(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "发货计划"
    for _ in range(4):
        ws.append([None] * 27)
    ws.append([
        "采购单号", "型号", "标签", "产品名称", "箱数", "箱容", "数量", "长", "宽", "高", "毛重",
        "CBM", "总材重", "总实重", "仓库", "FBA ID", "追踪编号", "ZD", "编号", "备注", "交货时间",
        "发货时间", "工厂", "DP", "货代", "出货单号", "状态", "so",
    ])
    # 一条历史记录：GH 工厂 + TD-RZ-419 型号，带箱容/长宽高/毛重，给"同工厂同型号"匹配用
    ws.append([
        "GH-2501009", "TD-RZ-419", "=+B6", "简约花瓶灰色树脂台灯", 60, 3, "=E6*F6", 550, 340, 440,
        13.6, None, None, None, "US", "FBA1", "TRACK1", "CA1", None, None, dt.datetime(2025, 3, 17),
        dt.datetime(2026, 1, 7), "GH", None, "KQ", "SK1", "已发货", None,
    ])
    ws.cell(row=6, column=4).fill = PatternFill("solid", fgColor="FFFF00")  # 产品名称格子的底色
    wb.save(path)


@pytest.fixture()
def tables(tmp_path):
    purchase_path = tmp_path / "purchase.xlsx"
    summary_path = tmp_path / "summary.xlsx"
    _write_purchase_summary(purchase_path)
    _write_shipment_summary(summary_path)
    return purchase_path, summary_path


def test_build_plan_matches_history_and_computes_boxes(tmp_path, tables):
    purchase_path, summary_path = tables
    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})

    assert not plan.skipped_orders
    assert not plan.skipped_files
    assert len(plan.items) == 1
    item = plan.items[0]
    assert item.supplier_code == "GH"
    assert item.box_capacity == 3
    assert item.length == 550 and item.width == 340 and item.height == 440
    assert item.gross_weight == 13.6
    assert item.boxes == 30  # 90 / 3，整除
    assert item.boxes_exact is True
    assert not item.notes


def test_build_plan_new_product_leaves_dims_blank_with_note(tmp_path, tables):
    purchase_path, summary_path = tables
    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[("TD-BRAND-NEW", "全新品", 100, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})

    item = plan.items[0]
    assert item.box_capacity is None
    assert item.length is None and item.width is None and item.height is None
    assert item.gross_weight is None
    assert item.boxes is None
    assert any("没有历史记录" in n for n in item.notes)


def test_build_plan_missing_supplier_mapping_leaves_code_blank(tmp_path, tables):
    purchase_path, summary_path = tables
    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="ZZ-001",
        supplier="没配过映射的供应商",
        rows=[("TD-X", "灯", 10, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {})

    item = plan.items[0]
    assert item.supplier_code is None
    assert any("映射代码" in n for n in item.notes)


def test_build_plan_skips_already_imported_order(tmp_path, tables):
    purchase_path, summary_path = tables
    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    # 这个订单号在采购汇总表里已经存在（GH-2501009）
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2501009",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "灯", 10, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})

    assert not plan.items
    assert len(plan.skipped_orders) == 1
    assert plan.skipped_orders[0].order_no == "GH-2501009"


def test_build_plan_indivisible_boxes_note(tmp_path, tables):
    purchase_path, summary_path = tables
    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609003",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "灯", 91, dt.datetime(2026, 10, 1))],  # 91 / 3 除不尽
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})

    item = plan.items[0]
    assert item.boxes_exact is False
    assert any("除不尽" in n for n in item.notes)


def test_apply_plan_appends_rows_to_both_sheets_with_seq_numbering(tmp_path, tables):
    purchase_path, summary_path = tables
    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[
            ("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1)),
            ("TD-NEW", "新品", 50, dt.datetime(2026, 10, 5)),
        ],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    apply_plan(plan, purchase_wb.active, summary_wb.active)

    p_ws = purchase_wb.active
    # 已有数据在第 4 行（表头第 2 行、子表头第 3 行），新行接着追加在第 5、6 行
    # 序号不是这张表自己按行独立编的号，是从订单号「GH-2609002」最后三位解析出来的（见
    # planner.py 的 _parse_seq_from_order_no），跟已有那一行的序号"001"没有递增关系，纯属
    # 巧合看起来像是"+1"——换一个订单号（比如"GH-2609099"）序号也会跟着变成"099"。
    assert p_ws.cell(row=5, column=1).value == "002"
    assert p_ws.cell(row=5, column=2).value == "GH-2609002"
    assert p_ws.cell(row=5, column=5).value == "GH"
    assert p_ws.cell(row=5, column=7).value == "TD-RZ-419"
    assert p_ws.cell(row=5, column=9).value == 90
    assert p_ws.cell(row=5, column=10).value == "pcs"
    assert p_ws.cell(row=5, column=6).value is None  # 店铺列模板行本来就是空的，抄过来还是空
    # 同一个订单号（GH-2609002）下的两个型号共用一个序号：值只写在第 5 行，第 6 行跟它合并
    # （合并单元格里非左上角的格子，openpyxl 读出来的 value 固定是 None，不能拿来断言）
    assert any(
        rng.min_row == 5 and rng.max_row == 6 and rng.min_col == 1 and rng.max_col == 1
        for rng in p_ws.merged_cells.ranges
    )
    # 未出货数量是公式，整行复制时应该按"自引用这一行"重新指向新行号（原来 =I4，新行是 =I5）
    assert p_ws.cell(row=5, column=12).value == "=I5"
    # 型号格子的底色格式也应该跟着抄过来
    assert p_ws.cell(row=5, column=7).fill.fgColor.rgb == p_ws.cell(row=4, column=7).fill.fgColor.rgb
    # 出货批次列（K 列，在数量单位和未出货数量之间）是模板行自己的出货历史，不该抄过来
    assert p_ws.cell(row=5, column=11).value is None

    s_ws = summary_wb.active
    # 已有数据在第 6 行，新行接着追加在第 7、8 行
    assert s_ws.cell(row=7, column=1).value == "GH-2609002"
    assert s_ws.cell(row=7, column=2).value == "TD-RZ-419"
    assert s_ws.cell(row=7, column=5).value == 30  # 箱数 = 90/3
    assert s_ws.cell(row=7, column=6).value == 3  # 箱容复制自历史
    assert s_ws.cell(row=7, column=8).value == 550  # 长
    assert s_ws.cell(row=7, column=22).value == "待定"  # 发货时间
    assert s_ws.cell(row=7, column=23).value == "GH"  # 工厂
    assert s_ws.cell(row=7, column=27).value == "未发货"  # 状态
    assert s_ws.cell(row=7, column=18).value is None  # ZD：模板行是"已发货"的旧值，不该抄过来
    assert s_ws.cell(row=7, column=3).value == "=+B7"  # 标签公式，自引用部分改指向新行
    assert s_ws.cell(row=7, column=7).value == "=E7*F7"  # 数量公式同理
    assert s_ws.cell(row=7, column=15).value is None  # 仓库：模板行是旧发货记录，不该抄过来
    assert s_ws.cell(row=7, column=16).value is None  # FBA ID 同理
    # 产品名称格子的底色格式也应该跟着抄过来
    assert s_ws.cell(row=7, column=4).fill.fgColor.rgb == s_ws.cell(row=6, column=4).fill.fgColor.rgb


def test_apply_plan_seq_comes_from_order_no_not_table_history_max(tmp_path, tables):
    # 回归测试：真实表格里「序号」不是这张表自己按行独立编的号，是订单号自带的信息——订单号
    # 最后三位数字就是该填的序号，业务方本来就是这么手填的。真实数据验证过：这张表的序号
    # 计数周期在中间某处（大概率是年份）重新起过，历史上出现过更早的行序号反而比更晚的行更
    # 大（比如 2025 年 12 月的一批到了 305，2026 年 9 月最新一行只有 213）——如果按"表格里
    # 全部历史序号取最大值 +1"，新订单会被写成一个跟订单号本身完全对不上、还离谱地偏大的号。
    # 这里在现有表格里插一条历史上"序号"畸高的行（模拟真实数据里 2025 年 12 月那批），确认
    # 新订单的序号只看自己订单号解出来的值，不受这条历史畸高行影响。
    purchase_path, summary_path = tables
    purchase_wb = openpyxl.load_workbook(purchase_path)
    p_ws = purchase_wb.active
    p_ws.append([
        "305", "WJ-2512305", dt.datetime(2025, 12, 1), dt.datetime(2026, 1, 1), "WJ", None,
        "TD-OLD-3", "历史畸高序号的旧型号", 30, "pcs", None, None, None, None, None,
    ])
    purchase_wb.save(purchase_path)

    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="SX-2609215",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    assert not plan.items[0].notes or "序号" not in "；".join(plan.items[0].notes)
    apply_plan(plan, purchase_wb.active, summary_wb.active)

    p_ws = purchase_wb.active
    # 新行接在第 6 行（第 4 行是原有数据，第 5 行是刚插的历史畸高序号行）；序号应该是订单号
    # 「SX-2609215」自己解出来的"215"，不是"305 + 1 = 306"。
    assert p_ws.cell(row=6, column=1).value == "215"


def test_apply_plan_seq_falls_back_to_table_max_when_order_no_unparseable(tmp_path, tables):
    # 订单号格式不是"XX-YYMMNNN"（解析不出最后三位当序号）的话，退回旧的"表格里已有的最大
    # 序号 +1"逻辑，不能什么都不写；build_plan 阶段应该在这一条记录的 notes 里提示，让人知道
    # 这个序号是"猜"出来的、需要核对，不是从订单号本身来的。
    purchase_path, summary_path = tables
    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="临时订单A",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    assert plan.items[0].seq_hint is None
    assert any("序号" in n for n in plan.items[0].notes)

    apply_plan(plan, purchase_wb.active, summary_wb.active)
    p_ws = purchase_wb.active
    # 表格里已有一行序号"001"，解析不出来就退回"最大值 +1"
    assert p_ws.cell(row=5, column=1).value == "002"


def test_apply_plan_copies_seq_column_style_when_template_row_is_a_merged_non_anchor_cell(tmp_path):
    # 回归测试：如果采购汇总表现有最后一行，恰好是某个多型号订单里"序号"列被纵向合并了的
    # 非首行（很常见——一个订单好几个型号，序号只写在第一行、其余行合并），openpyxl 会把
    # 这种非左上角的合并格子变成 MergedCell，读它的样式永远是"没有样式"（has_style 恒为
    # False）。之前 copy_row() 直接读模板行这一格的样式，读到的就是这种"假的没有样式"，
    # 抄出来的新行「序号」格子会丢掉字体加粗、居中这些格式。现在应该去读这个合并区域左上角
    # 那个格子的真实样式。
    from openpyxl.styles import Alignment, Font

    purchase_path = tmp_path / "purchase.xlsx"
    summary_path = tmp_path / "summary.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "采购订单汇总"
    ws.append([None] * 12)
    ws.append([
        "序号", "订单号", "采购日期", "交货日期", "供应商名称", "店铺", "型号", "产品名称",
        "订单数量", "数量单位", "出货日期占位", "未出货数量",
    ])
    ws.append([None] * 10 + ["出货时间", None])
    # 现有订单占了两行（两个型号），序号列纵向合并，值只写在第 4 行（左上角）
    ws.append([
        "001", "GH-2501001", dt.datetime(2025, 1, 7), dt.datetime(2025, 3, 17), "GH", None,
        "TD-OLD-1", "旧型号1", 90, "pcs", None, "=I4",
    ])
    ws.append([
        None, "GH-2501001", dt.datetime(2025, 1, 7), dt.datetime(2025, 3, 17), "GH", None,
        "TD-OLD-2", "旧型号2", 60, "pcs", None, "=I5",
    ])
    ws.cell(row=4, column=1).font = Font(bold=True)
    ws.cell(row=4, column=1).alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=4, start_column=1, end_row=5, end_column=1)
    wb.save(purchase_path)
    _write_shipment_summary(summary_path)

    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    apply_plan(plan, purchase_wb.active, summary_wb.active)

    p_ws = purchase_wb.active
    # 新行接在第 6 行，模板行取的是第 5 行——那正是上面合并区域里的非左上角行
    new_seq_cell = p_ws.cell(row=6, column=1)
    assert new_seq_cell.value == "002"
    assert new_seq_cell.font.bold is True
    assert new_seq_cell.alignment.horizontal == "center"


def test_apply_plan_unmerges_stray_merge_that_overlaps_new_rows(tmp_path, tables):
    # 回归测试：真实表格里常见手工把"序号"列提前合并了一大片还没用到的空白行（比如一路
    # 合并到第 50 行，方便以后陆续往下填）。追加的新行如果正好落进这种旧合并区域，那一格
    # 会是 openpyxl 的 MergedCell（合并区域里非左上角的格子），直接给它赋值会抛
    # AttributeError("MergedCell object attribute 'value' is read-only")——之前没处理这种
    # 情况，现在应该先把跟新行重叠的这部分旧合并拆开，再正常写入。
    purchase_path, summary_path = tables
    purchase_wb = openpyxl.load_workbook(purchase_path)
    p_ws = purchase_wb.active
    # 现有数据只到第 4 行，"序号"列却提前合并到了第 50 行（跟真实数据没关系，纯粹是提前
    # 格式化的空白行）
    p_ws.merge_cells(start_row=4, start_column=1, end_row=50, end_column=1)
    purchase_wb.save(purchase_path)

    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    apply_plan(plan, purchase_wb.active, summary_wb.active)  # 不应该抛异常

    p_ws = purchase_wb.active
    assert p_ws.cell(row=5, column=1).value == "002"


def test_apply_plan_preserves_dim_formulas_in_shipment_summary(tmp_path):
    # 发货计划汇总表的最后一行（模板行）「长/宽/高」是公式（跟着箱容自动算），不是写死数字——
    # 新增行应该保留这个公式（自引用部分重指向新行），不能被写死的历史数字覆盖掉。
    purchase_path = tmp_path / "purchase.xlsx"
    summary_path = tmp_path / "summary.xlsx"
    _write_purchase_summary(purchase_path)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "发货计划"
    for _ in range(4):
        ws.append([None] * 27)
    ws.append([
        "采购单号", "型号", "标签", "产品名称", "箱数", "箱容", "数量", "长", "宽", "高", "毛重",
        "CBM", "总材重", "总实重", "仓库", "FBA ID", "追踪编号", "ZD", "编号", "备注", "交货时间",
        "发货时间", "工厂", "DP", "货代", "出货单号", "状态", "so",
    ])
    ws.append([
        "GH-2501009", "TD-RZ-419", "=+B6", "简约花瓶灰色树脂台灯", 60, 3, "=E6*F6",
        "=F6*10", "=F6*20", "=F6*30",
        13.6, None, None, None, "US", "FBA1", "TRACK1", "CA1", None, None, dt.datetime(2025, 3, 17),
        dt.datetime(2026, 1, 7), "GH", None, "KQ", "SK1", "已发货", None,
    ])
    wb.save(summary_path)

    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    apply_plan(plan, purchase_wb.active, summary_wb.active)

    s_ws = summary_wb.active
    # 新行接在第 7 行，长/宽/高应该还是公式，自引用的 F6 改成了 F7
    assert s_ws.cell(row=7, column=8).value == "=F7*10"
    assert s_ws.cell(row=7, column=9).value == "=F7*20"
    assert s_ws.cell(row=7, column=10).value == "=F7*30"


def test_apply_plan_preserves_dim_formulas_for_duplicate_named_columns(tmp_path):
    # 回归测试：真实表里出现过"长/宽/高"这几个表头重复出现不止一次（历史遗留的重复列）——
    # 只处理第一次出现的位置的话，第二次出现那一列会被 copy_row() 原样抄一份模板行当时
    # 写死的数字，不会跟着箱型自动变化了。这里第二组"长/宽/高"（列 29/30/31）模板行是写死的
    # 数字，往上一行（第 6 行）才是公式——新行应该往上找到那一行的公式抄一份，不是简单複製
    # 模板行的写死数字。
    purchase_path = tmp_path / "purchase.xlsx"
    summary_path = tmp_path / "summary.xlsx"
    _write_purchase_summary(purchase_path)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "发货计划"
    for _ in range(4):
        ws.append([None] * 31)
    ws.append([
        "采购单号", "型号", "标签", "产品名称", "箱数", "箱容", "数量", "长", "宽", "高", "毛重",
        "CBM", "总材重", "总实重", "仓库", "FBA ID", "追踪编号", "ZD", "编号", "备注", "交货时间",
        "发货时间", "工厂", "DP", "货代", "出货单号", "状态", "so",
        "长", "宽", "高",  # 重复出现的第二组长宽高（列 29/30/31）
    ])
    ws.append([
        "GH-2501001", "TD-RZ-000", "=+B6", "旧型号灯饰", 60, 3, "=E6*F6",
        "=F6*10", "=F6*20", "=F6*30",
        13.6, None, None, None, "US", "FBA0", "TRACK0", "CA1", None, None, dt.datetime(2025, 1, 1),
        dt.datetime(2025, 6, 1), "GH", None, "KQ", "SK0", "已发货", None,
        "=F6*10", "=F6*20", "=F6*30",
    ])
    ws.append([
        "GH-2501009", "TD-RZ-419", "=+B7", "简约花瓶灰色树脂台灯", 60, 3, "=E7*F7",
        "=F7*10", "=F7*20", "=F7*30",
        13.6, None, None, None, "US", "FBA1", "TRACK1", "CA1", None, None, dt.datetime(2025, 3, 17),
        dt.datetime(2026, 1, 7), "GH", None, "KQ", "SK1", "已发货", None,
        # 模板行（最后一行）的第二组长宽高是历史手填的写死数字，不是公式
        600, 1200, 1800,
    ])
    wb.save(summary_path)

    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    apply_plan(plan, purchase_wb.active, summary_wb.active)

    s_ws = summary_wb.active
    # 新行接在第 8 行。第一组长/宽/高（列 8/9/10）模板行本来就是公式，照旧重指向新行。
    assert s_ws.cell(row=8, column=8).value == "=F8*10"
    assert s_ws.cell(row=8, column=9).value == "=F8*20"
    assert s_ws.cell(row=8, column=10).value == "=F8*30"
    # 第二组长/宽/高（列 29/30/31）模板行是写死数字，应该往上找到第 6 行的公式抄一份、
    # 重指向新行，而不是被 copy_row() 原样抄一份模板行写死的 600/1200/1800。
    assert s_ws.cell(row=8, column=29).value == "=F8*10"
    assert s_ws.cell(row=8, column=30).value == "=F8*20"
    assert s_ws.cell(row=8, column=31).value == "=F8*30"


def test_apply_plan_reindexes_array_formula_for_second_dim_group(tmp_path):
    # 回归测试：真实表里第二组"长/宽/高"是一个 Excel 传统数组公式（一次 XLOOKUP 覆盖
    # 长/宽/高三个格子，比如 =XLOOKUP(...,...,$T:$V)）——公式原文只存在"长"这一格
    # （openpyxl 读出来是 ArrayFormula 对象），"宽"/"高"在文件里读出来只是缓存的普通数字，
    # 永远不会有公式文本。之前的逻辑把 ArrayFormula 当成"不是公式"，"长"往上找也只会找到
    # 别的 ArrayFormula（同样不被识别），最后退回写死数字；"宽"/"高"更是天生就没有独立公式
    # 可找，永远退回写死数字——三列全变成跟第一组长宽高一样的历史写死数字。现在应该正确
    # 识别 ArrayFormula，把"长"重新指向新行（ref 和公式原文里的自引用都要改），"宽"/"高"
    # 留空交给 Excel 重新计算。
    purchase_path = tmp_path / "purchase.xlsx"
    summary_path = tmp_path / "summary.xlsx"
    _write_purchase_summary(purchase_path)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "发货计划"
    for _ in range(4):
        ws.append([None] * 31)
    ws.append([
        "采购单号", "型号", "标签", "产品名称", "箱数", "箱容", "数量", "长", "宽", "高", "毛重",
        "CBM", "总材重", "总实重", "仓库", "FBA ID", "追踪编号", "ZD", "编号", "备注", "交货时间",
        "发货时间", "工厂", "DP", "货代", "出货单号", "状态", "so",
        "长", "宽", "高",  # 重复出现的第二组长宽高（列 29/30/31），真实表里是数组公式
    ])
    ws.append([
        "GH-2501009", "TD-RZ-419", "=+B6", "简约花瓶灰色树脂台灯", 60, 3, "=E6*F6",
        "=F6*10", "=F6*20", "=F6*30",
        13.6, None, None, None, "US", "FBA1", "TRACK1", "CA1", None, None, dt.datetime(2025, 3, 17),
        dt.datetime(2026, 1, 7), "GH", None, "KQ", "SK1", "已发货", None,
        None, None, None,  # 先占位，下面用 ArrayFormula 单独写"长"这一格
    ])
    ws.cell(row=6, column=29, value=ArrayFormula(
        ref="AC6:AE6",
        text="=_xlfn.XLOOKUP(W6&B6,[1]在售产品信息总表!$M:$M&[1]在售产品信息总表!$H:$H,[1]在售产品信息总表!$T:$V)",
    ))
    ws.cell(row=6, column=30, value=340)  # "宽"：数组公式溢出的缓存数字，不是公式
    ws.cell(row=6, column=31, value=440)  # "高"：同上
    wb.save(summary_path)

    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    apply_plan(plan, purchase_wb.active, summary_wb.active)

    s_ws = summary_wb.active
    # 新行接在第 7 行。"长"应该还是数组公式，ref 和公式里的自引用都重指向第 7 行。
    new_length = s_ws.cell(row=7, column=29).value
    assert isinstance(new_length, ArrayFormula)
    assert new_length.ref == "AC7:AE7"
    assert new_length.text == (
        "=_xlfn.XLOOKUP(W7&B7,[1]在售产品信息总表!$M:$M&[1]在售产品信息总表!$H:$H,"
        "[1]在售产品信息总表!$T:$V)"
    )
    # "宽"/"高"是这个数组公式的溢出结果，不该写死成历史数字（600/1200 那种），留空交给
    # Excel 打开时自己重新算。
    assert s_ws.cell(row=7, column=30).value is None
    assert s_ws.cell(row=7, column=31).value is None


def test_apply_plan_ignores_broken_template_row_and_always_uses_first_data_row(tmp_path):
    # 回归测试：真实表里出现过"长/宽/高"历史上没有统一套用公式——早期数据是手填的写死
    # 数字，中间某次改版才开始套公式，且中途偶尔会混进个别残缺/异常的行（公式跟别的行对不
    # 上）。如果按"以当前模板行为准、找不到再往上找最近一行"来处理，一旦最靠近新行的那一行
    # （模板行本身，或离它最近的一行）刚好是这种异常数据，会悄悄拷贝出错误结果。现在应该
    # 固定认表格第一条数据行的表达式，不管中间/模板行数据再乱都不受影响。
    purchase_path = tmp_path / "purchase.xlsx"
    summary_path = tmp_path / "summary.xlsx"
    _write_purchase_summary(purchase_path)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "发货计划"
    for _ in range(4):
        ws.append([None] * 28)
    ws.append([
        "采购单号", "型号", "标签", "产品名称", "箱数", "箱容", "数量", "长", "宽", "高", "毛重",
        "CBM", "总材重", "总实重", "仓库", "FBA ID", "追踪编号", "ZD", "编号", "备注", "交货时间",
        "发货时间", "工厂", "DP", "货代", "出货单号", "状态", "so",
    ])
    ws.append([
        "GH-2501001", "TD-RZ-000", "=+B6", "第一条数据行", 60, 3, "=E6*F6",
        "=F6*10", "=F6*20", "=F6*30",  # 表格第一条数据行——正确的表达式，应该被固定沿用
        13.6, None, None, None, "US", "FBA0", "TRACK0", "CA1", None, None, dt.datetime(2025, 1, 1),
        dt.datetime(2025, 6, 1), "GH", None, "KQ", "SK0", "已发货", None,
    ])
    ws.append([
        "GH-2501009", "TD-RZ-419", "=+B7", "简约花瓶灰色树脂台灯", 60, 3, "=E7*F7",
        # 模板行（最后一行）是残缺/异常数据：只有"长"是公式，"宽"/"高"是明显对不上箱数的
        # 写死数字（模拟历史上某次手动改坏了这一行）——不应该被拿来当模板
        "=F7*999", 12345, 67890,
        13.6, None, None, None, "US", "FBA1", "TRACK1", "CA1", None, None, dt.datetime(2025, 3, 17),
        dt.datetime(2026, 1, 7), "GH", None, "KQ", "SK1", "已发货", None,
    ])
    wb.save(summary_path)

    order_folder = tmp_path / "orders"
    order_folder.mkdir()
    _write_order_file(
        order_folder / "order1.xlsx",
        order_no="GH-2609002",
        supplier="广东GH工厂",
        rows=[("TD-RZ-419", "简约花瓶灰色树脂台灯", 90, dt.datetime(2026, 10, 1))],
    )

    purchase_wb = openpyxl.load_workbook(purchase_path)
    summary_wb = openpyxl.load_workbook(summary_path)
    plan = build_plan(order_folder, purchase_wb.active, summary_wb.active, {"广东GH工厂": "GH"})
    apply_plan(plan, purchase_wb.active, summary_wb.active)

    s_ws = summary_wb.active
    # 新行接在第 8 行——长/宽/高应该固定沿用第一条数据行（第 6 行）的表达式，不受模板行
    # （第 7 行，残缺异常）影响。
    assert s_ws.cell(row=8, column=8).value == "=F8*10"
    assert s_ws.cell(row=8, column=9).value == "=F8*20"
    assert s_ws.cell(row=8, column=10).value == "=F8*30"
