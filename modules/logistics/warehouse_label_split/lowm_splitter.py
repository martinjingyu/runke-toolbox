"""拆分 LO-WM（Walmart）站点的箱唛 PDF：不需要认 SKU，直接按页分配给各厂商。

这类站点 PDF（比如 DFW5s.pdf）本身是纯图片箱唛，没有可提取的文字，拿不到（也不需要）每一页
具体是哪个 SKU——用户确认过 Walmart 这个平台只在乎"这一批货的总箱数"，不在乎每箱具体装的是
什么，各厂商之间也没有必须遵守的页面顺序要求。所以不用像 CA1/CG 拆分工具那样逐页解析 SKU
再回头核对，直接按发货计划表里这个仓库、状态=未发货的记录按「工厂」把「箱数」汇总起来，
按顺序把这么多页从 PDF 里切给这个厂商就够了。

流程：
    1. 站点代号从文件名推导（复用 splitter.py 的 derive_warehouse_code）。
    2. 发货计划表里，找「仓库含站点代号 + 状态=未发货」的行，按「工厂」汇总「箱数」
       （见 shipping_plan.py 的 load_pending_boxes_by_factory）。
    3. 按工厂代号排序（没有顺序要求，选一个固定、可预测的顺序，跟其它拆分工具的排序习惯
       一致），依次把这么多页从 PDF 里切下来给这个工厂，页码紧跟着上一个工厂切完的位置继续
       切，不重叠、不跳页。
    4. 总需求箱数如果比 PDF 总页数还多，切到哪个厂商发现页不够了就停：不给这个厂商生成文件，
       后面排在它之后的厂商也不再处理（页数从这里开始就已经对不上了，继续往下切没有意义），
       写清楚缺口交给人工核对，不猜着多切/少切。总需求比总页数少，剩下没分完的页也不生成
       文件，写清楚剩了多少页，同样交给人工看（可能是这份 PDF 混了别的仓库的箱子，或者发货
       计划表这边总数没登记全）。
    5. 输出文件放进"工厂/仓库"两层文件夹（跟 CA1/CG 拆分工具一样的结构），文件名是
       "工厂 仓库 箱数箱.pdf"。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import fitz

from .shipping_plan import load_pending_boxes_by_factory
from .splitter import derive_warehouse_code


@dataclass
class FactoryAllocation:
    factory: str
    boxes: int
    output_path: Path


@dataclass
class LowmSplitReport:
    outputs: list[FactoryAllocation] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def run(label_pdf_path: str | Path, shipping_plan_path: str | Path) -> LowmSplitReport:
    label_pdf_path = Path(label_pdf_path)
    warehouse_code = derive_warehouse_code(label_pdf_path)

    doc = fitz.open(label_pdf_path)
    try:
        totals = load_pending_boxes_by_factory(shipping_plan_path, warehouse_code)

        report = LowmSplitReport()
        if not totals:
            report.notes.append(f"发货计划表里查不到仓库含 {warehouse_code} + 未发货的记录，未拆分")
            return report

        page_count = doc.page_count
        cursor = 0
        stopped_early = False
        for factory in sorted(totals):
            total = totals[factory]
            if not float(total).is_integer():
                report.notes.append(f"「{factory}」：箱数合计是 {total}，不是整数，没法据此切页，未拆分，需要人工核对")
                stopped_early = True
                break
            boxes = int(total)
            if boxes <= 0:
                continue

            if cursor + boxes > page_count:
                remaining = page_count - cursor
                report.notes.append(
                    f"「{factory}」：需要 {boxes} 箱，但 PDF 只剩 {remaining} 页了，页数不够，"
                    f"从这里开始都没有再切分，需要人工核对"
                )
                stopped_early = True
                break

            out_doc = fitz.open()
            out_doc.insert_pdf(doc, from_page=cursor, to_page=cursor + boxes - 1)

            out_dir = label_pdf_path.parent / _sanitize(factory) / _sanitize(warehouse_code)
            out_dir.mkdir(parents=True, exist_ok=True)
            out_name = f"{_sanitize(factory)} {_sanitize(warehouse_code)} {boxes}箱.pdf"
            out_path = out_dir / out_name
            out_doc.save(out_path)
            out_doc.close()

            report.outputs.append(FactoryAllocation(factory=factory, boxes=boxes, output_path=out_path))
            cursor += boxes

        if not stopped_early and cursor < page_count:
            report.notes.append(f"PDF 还剩 {page_count - cursor} 页没有分配出去，发货计划表这边这几个厂商的箱数合计没用完整份 PDF，需要人工核对")

        return report
    finally:
        doc.close()


def _sanitize(name: str) -> str:
    cleaned = "".join(c if c.isalnum() or c in "-_." else "_" for c in name)
    return cleaned or "unknown"
