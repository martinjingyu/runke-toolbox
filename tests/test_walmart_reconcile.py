from pathlib import Path

import openpyxl
import pytest

from modules.logistics.walmart_shipment_reconcile.box_label import _HAS_TESSERACT, ocr_single_sku
from modules.logistics.walmart_shipment_reconcile.reconcile import _sanitize_filename, run
from modules.logistics.walmart_shipment_reconcile.shipping_plan import normalize_warehouse, parse_shipping_plan
from modules.logistics.walmart_shipment_reconcile.translation import normalize_sku, parse_translation

REAL_DATA_DIR = Path("/Users/jingyuhuang/Documents/Work/闰科/物流仓库/发货核对")
REAL_PDF = REAL_DATA_DIR / "DFW5.pdf"
REAL_TRANSLATION = REAL_DATA_DIR / "测试8.24.xlsx"
REAL_PLAN = REAL_DATA_DIR / "发货计划表.xlsx 5.18 .xlsx1.xlsx"


def test_sanitize_filename_strips_unsafe_chars():
    assert _sanitize_filename("WMTD-256") == "WMTD-256"
    assert _sanitize_filename("a/b\\c:d") == "a_b_c_d"
    assert _sanitize_filename("") == "unknown"


def test_normalize_sku_matches_on_letter_and_digit_runs_only():
    # 分隔符（有没有、是什么符号）完全不影响归一化结果，只看字母串/数字串本身的内容和顺序
    assert normalize_sku("wmtd — 256") == "WMTD-256"
    assert normalize_sku("WMTD202") == "WMTD-202"
    assert normalize_sku("WMTD-202") == "WMTD-202"
    assert normalize_sku("WMTD—-617") == "WMTD-617"  # OCR 常见的错读：一个"-"读成两个字符
    assert normalize_sku("WM-CK-FL-1112") == "WM-CK-FL-1112"
    assert normalize_sku(None) == ""


def test_normalize_warehouse_extracts_code_from_various_formats():
    assert normalize_warehouse("US(IND3)") == "IND3"
    assert normalize_warehouse("PHL5s") == "PHL5"
    assert normalize_warehouse("DFW5s") == "DFW5"
    assert normalize_warehouse("CG-RQ-MX") is None  # 不是仓库代码格式，识别不出来很正常


def test_parse_translation_reads_sku_and_original_sku_columns(tmp_path):
    path = tmp_path / "translation.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([" SKU", "原sku", "工厂", "仓库", "求和项:箱数", "求和项: 已发货商品"])
    ws.append(["WMTD-256", "TD-392", "SX", "DFW5s", 1, 3])
    wb.save(path)

    mapping = parse_translation(path)
    assert mapping[normalize_sku("WMTD-256")] == "TD-392"


def test_parse_shipping_plan_finds_header_row_and_reads_tracking_id(tmp_path):
    path = tmp_path / "plan.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["出货单号"])
    ws.append(["出货日期"])
    ws.append(["采购单号", "型号", "标签", "产品名称", "箱数", "箱容", "数量", "仓库", "追踪编号"])
    ws.append(["GH-001", "TD-RZ-419", "TD-392", "台灯", 1, 3, 3, "US(DFW5s)", "1015658WFB"])
    wb.save(path)

    rows = parse_shipping_plan(path)
    assert len(rows) == 1
    assert rows[0].sku == "TD-392"
    assert rows[0].warehouse_code == "DFW5"
    assert rows[0].tracking_id == "1015658WFB"
    assert rows[0].planned_quantity == 3


@pytest.mark.slow
@pytest.mark.skipif(not REAL_PDF.exists(), reason="需要本机真实的物流仓库数据文件")
@pytest.mark.skipif(not _HAS_TESSERACT, reason="需要装 Tesseract（brew install tesseract）")
def test_ocr_single_sku_reads_known_label():
    # DFW5.pdf 第 0 页手工核对过箱唛印刷文字是 WMTD-256。ocr_single_sku 只负责去掉空白噪音，
    # 不做大小写/破折号归一化（那是 normalize_sku 的活），所以这里要经过 normalize_sku 再比较，
    # 跟 run() 内部的实际用法一致。
    raw = ocr_single_sku(REAL_PDF, 0)
    assert normalize_sku(raw) == "WMTD-256"


@pytest.mark.slow
@pytest.mark.skipif(
    not (REAL_PDF.exists() and REAL_TRANSLATION.exists() and REAL_PLAN.exists()),
    reason="需要本机真实的物流仓库数据文件",
)
@pytest.mark.skipif(not _HAS_TESSERACT, reason="需要装 Tesseract（brew install tesseract）")
def test_run_against_real_dfw5_data(tmp_path):
    report = run([REAL_PDF], REAL_TRANSLATION, REAL_PLAN, tmp_path)

    # 47 个箱标签页 + 3 个托盘汇总页（"PALLET X OF 3"，没有 GTIN，应该被跳过而不是报错）
    assert report.skipped_pages["DFW5.pdf"] == 3
    assert sum(r.box_count for r in report.results) == 47

    # 拆出来的 PDF 都是真实存在的文件
    assert report.split_pdf_paths
    for path in report.split_pdf_paths:
        assert path.exists()

    # 手工核对过的一条完整链路：GTIN 00763073694506 -> OCR 出 WMTD-256 -> 翻译表查出货号 TD-392
    # -> 发货计划表里 shipment 1015658WFB + TD-392 计划数量是 3，箱唛上也正好是 1 箱 3 个，应该一致
    known = next(r for r in report.results if r.shipment_id == "1015658WFB" and r.sku == "TD-392")
    assert known.planned_quantity == 3
    assert known.actual_quantity == 3
    assert known.box_count == 1
    assert known.match is True
