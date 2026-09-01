from collections import Counter
from pathlib import Path

import fitz
import pytest

from modules.logistics.fba_label_redact.redact import find_pdfs, redact_page, run

REAL_INPUT_DIR = Path(
    "/Users/jingyuhuang/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/"
    "wxid_xpgmswv2dt5q22_c33f/msg/file/2026-09/FBA标/FBA标原文档"
)


def _make_label_page(doc: fitz.Document, *, fc_code: str, country: str, sender: str = "somesenderpinyin") -> None:
    """按真实箱唛的结构（目的地5行/发货地4行）造一个最小可用的测试页面。"""
    page = doc.new_page(width=306, height=200)
    x0, ox0 = 17.3, 154.0
    page.insert_text((x0, 8), "目的地：", fontname="china-s", fontsize=8)
    page.insert_text((x0, 16), f"FBA: {sender}", fontname="Helvetica", fontsize=6)
    page.insert_text((x0, 24), fc_code, fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 32), "1 MAIN ST", fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 40), "SOMEWHERE, ST 00000", fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 48), country, fontname="china-s", fontsize=8)

    page.insert_text((ox0, 8), "发货地：", fontname="china-s", fontsize=8)
    page.insert_text((ox0, 16), sender, fontname="Helvetica", fontsize=7)
    page.insert_text((ox0, 24), "Guangdong - somecity - 000000", fontname="Helvetica", fontsize=7)
    page.insert_text((ox0, 32), "some street 1", fontname="Helvetica", fontsize=8)
    page.insert_text((ox0, 40), "中国", fontname="china-s", fontsize=8)


def test_find_pdfs_only_matches_pdf_files(tmp_path):
    (tmp_path / "a.pdf").write_bytes(b"")
    (tmp_path / "b.PDF").write_bytes(b"")  # 大写扩展名不算，跟真实文件夹里都是小写 .pdf 一致
    (tmp_path / "notes.txt").write_bytes(b"")
    found = find_pdfs(tmp_path)
    assert [p.name for p in found] == ["a.pdf"]


def test_redact_page_unknown_country_is_skipped_not_guessed():
    doc = fitz.open()
    _make_label_page(doc, fc_code="ABC1", country="墨西哥")
    result = redact_page(doc[0], "test.pdf", 0)
    assert result.modified is False
    assert "墨西哥" in result.status


def test_redact_page_us_keeps_origin_block_trimmed():
    doc = fitz.open()
    _make_label_page(doc, fc_code="IND9", country="美国")
    result = redact_page(doc[0], "test.pdf", 0)
    assert result.modified is True
    assert result.status == "美国"

    text = doc[0].get_text()
    assert "somesenderpinyin" not in text  # 发货人名字必须被删掉
    assert "FBA: IND9" in text  # FC 代码跟 FBA: 合并到一行
    assert "发货地：" in text  # 美国的话，发货地这个标签还在
    assert "Guangdong" in text  # 发货地剩下的内容还在


def test_redact_page_canada_removes_whole_origin_block():
    doc = fitz.open()
    _make_label_page(doc, fc_code="YEG2", country="加拿大")
    result = redact_page(doc[0], "test.pdf", 0)
    assert result.modified is True
    assert result.status == "加拿大"

    text = doc[0].get_text()
    assert "somesenderpinyin" not in text
    assert "FBA: YEG2" in text
    assert "发货地" not in text  # 加拿大：连"发货地"这个标签本身都要没了
    assert "Guangdong" not in text


def test_redact_page_skips_when_dest_first_line_not_fba_prefixed():
    # 光按行数（5行/4行）判断不够——如果目的地第一行不是"FBA:"开头，说明这一页的结构
    # 跟预期的不一样（可能是别的字段占了这个位置），不该照样把它当发货人名字删掉
    doc = fitz.open()
    page = doc.new_page(width=306, height=200)
    x0 = 17.3
    page.insert_text((x0, 8), "目的地：", fontname="china-s", fontsize=8)
    page.insert_text((x0, 16), "NOT A SENDER LINE", fontname="Helvetica", fontsize=6)
    page.insert_text((x0, 24), "IND9", fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 32), "1 MAIN ST", fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 40), "SOMEWHERE, ST 00000", fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 48), "美国", fontname="china-s", fontsize=8)
    ox0 = 154.0
    page.insert_text((ox0, 8), "发货地：", fontname="china-s", fontsize=8)
    page.insert_text((ox0, 16), "somesenderpinyin", fontname="Helvetica", fontsize=7)
    page.insert_text((ox0, 24), "Guangdong - somecity - 000000", fontname="Helvetica", fontsize=7)
    page.insert_text((ox0, 32), "some street 1", fontname="Helvetica", fontsize=8)
    page.insert_text((ox0, 40), "中国", fontname="china-s", fontsize=8)

    result = redact_page(doc[0], "test.pdf", 0)
    assert result.modified is False
    assert "FBA:" in result.status
    # 没处理过，原始内容应该原封不动还在
    assert "NOT A SENDER LINE" in doc[0].get_text()


def test_redact_page_does_not_require_dest_and_origin_sender_names_to_match():
    # 不同发货人名字的写法本来就可能不完全一样（比如目的地那边有个后缀，发货地没有），
    # 不该强行要求两边文字一字不差——否则会把本该正常处理的页面也拦下来。
    # 目的地写 senderA，发货地写完全不一样的 senderB，这一页应该照样能处理成功。
    doc = fitz.open()
    page = doc.new_page(width=306, height=200)
    x0, ox0 = 17.3, 154.0
    page.insert_text((x0, 8), "目的地：", fontname="china-s", fontsize=8)
    page.insert_text((x0, 16), "FBA: senderA", fontname="Helvetica", fontsize=6)
    page.insert_text((x0, 24), "IND9", fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 32), "1 MAIN ST", fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 40), "SOMEWHERE, ST 00000", fontname="Helvetica", fontsize=8)
    page.insert_text((x0, 48), "美国", fontname="china-s", fontsize=8)
    page.insert_text((ox0, 8), "发货地：", fontname="china-s", fontsize=8)
    page.insert_text((ox0, 16), "senderB", fontname="Helvetica", fontsize=7)
    page.insert_text((ox0, 24), "Guangdong - somecity - 000000", fontname="Helvetica", fontsize=7)
    page.insert_text((ox0, 32), "some street 1", fontname="Helvetica", fontsize=8)
    page.insert_text((ox0, 40), "中国", fontname="china-s", fontsize=8)

    result = redact_page(doc[0], "test.pdf", 0)
    assert result.modified is True
    assert result.status == "美国"
    text = doc[0].get_text()
    assert "senderA" not in text
    assert "senderB" not in text
    assert "FBA: IND9" in text


@pytest.mark.slow
@pytest.mark.skipif(not REAL_INPUT_DIR.exists(), reason="需要本机真实的 FBA 标签样例文件")
def test_run_against_real_reference_files(tmp_path):
    report = run(REAL_INPUT_DIR, tmp_path)

    # 手工核对过的真实数据：6 个文件一共 85 页，70 页美国 + 15 页加拿大，全部能正常处理，
    # 没有跳过的（结构在这 6 个文件里是完全一致的）
    assert len(report.results) == 85
    assert report.modified_count == 85
    assert report.skipped_results == []
    assert Counter(r.status for r in report.results) == Counter(美国=70, 加拿大=15)
    assert len(report.output_paths) == 6
