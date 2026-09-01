"""命令行跑一次"Walmart 发货数量核对"，方便批处理/调试用（图形界面里也能跑这个工具）。

用法：
    python scripts/run_walmart_reconcile.py 翻译表.xlsx 发货计划表.xlsx 箱唛1.pdf 箱唛2.pdf ... --out 输出目录

注意：这个工具自己的依赖（openpyxl/pymupdf/pylibdmtx/pytesseract + Tesseract）不在项目的
基础 requirements.txt 里——图形界面走的是"点开这个工具才检查/安装"，命令行这条路没有界面
弹窗提示，所以依赖没装齐的话会直接从 run() 里报错，需要自己先装好（见 SETUP.md）。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from modules.logistics.walmart_shipment_reconcile.reconcile import run, write_report_xlsx


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("translation", type=Path, help="翻译表（WM-SKU -> 货号），比如 测试8.24.xlsx")
    parser.add_argument("plan", type=Path, help="发货计划表，比如 发货计划表.xlsx 5.18 .xlsx1.xlsx")
    parser.add_argument("pdfs", type=Path, nargs="+", help="箱唛 PDF，一个仓一个文件")
    parser.add_argument("--out", type=Path, default=Path("output"), help="输出目录，默认 ./output")
    args = parser.parse_args()

    report = run(args.pdfs, args.translation, args.plan, args.out)

    report_path = args.out / "核对结果.xlsx"
    write_report_xlsx(report, report_path)

    print(f"拆出的箱唛 PDF：{len(report.split_pdf_paths)} 个，存在 {args.out}")
    for name, skipped in report.skipped_pages.items():
        if skipped:
            print(f"  {name}：跳过了 {skipped} 页非箱标签页（比如托盘汇总页）")

    if report.unresolved_gtins:
        print(f"有 {len(report.unresolved_gtins)} 个 GTIN 没能读出/翻译出 SKU，拆分文件用的是能读到的原始文字或 GTIN 命名：")
        for gtin in sorted(report.unresolved_gtins):
            print(f"  {gtin}")

    mismatched = [r for r in report.results if not r.match]
    print(f"核对结果表：{report_path}")
    print(f"共 {len(report.results)} 条（SHIPMENT ID + 货号），其中 {len(mismatched)} 条不一致")
    for r in mismatched:
        plan_note = "" if r.in_plan else "（计划表里没查到）"
        print(f"  [不一致] {r.sku} {r.warehouse} shipment={r.shipment_id} 计划={r.planned_quantity} 实际={r.actual_quantity} {plan_note}")


if __name__ == "__main__":
    main()
