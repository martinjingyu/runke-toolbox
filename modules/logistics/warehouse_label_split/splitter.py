"""拆分入口：先看标签 PDF（比如 CA1.pdf、CG-TS-MD.pdf）里实际出现过哪些标签，再拿这份名单去
发货计划表里逐个核对——只查这些标签，不读发货计划表里其它几万行跟这份 PDF 无关的历史数据
（那些标签根本不在这份 PDF 里，读出来也没用，还会在报告里制造一堆无关的噪音）。核对到「仓库
含站点代号、状态=未发货」的记录（不要求发货时间=待定，已经安排了具体发货日期但还没实际发出
去的也算），就把对应的箱子页面从标签 PDF 里抽出来，存到标签 PDF 所在目录下新建的
"厂商代号/站点代号"两层子文件夹（比如"GH/CA1"），文件名按"标签 箱数合计箱.pdf"命名（比如
"TD-243 10箱.pdf"）——厂商已经体现在文件夹名字里了，文件名不用重复写一遍。不同厂商的标签
分到不同的厂商文件夹，同一个厂商的多个标签落进同一个厂商文件夹下面同一个站点子文件夹。

站点代号（拿去筛发货计划表「仓库」列的关键字）不写死，从标签 PDF 的文件名推导：取文件名
（不含扩展名）里第一个"-"之前的部分，没有"-"就用整个文件名——"CA1.pdf"和"CA1-XXX.pdf"都
推出"CA1"，"CG-TS-MD.pdf"推出"CG"，跟目前见过的两个站点的实际命名习惯都对得上。不同站点
的标签 PDF 版式差别很大（CA1 是"一个箱子两页"，CG 是"一个箱子一页"，见 label_pdf.py /
label_pdf_cg.py），但版式差异只影响"怎么从 PDF 里解析出箱子列表"这一步，所以那部分做成
调用方传入的 parse_boxes 函数，剩下匹配发货计划表、拆分、命名的逻辑都是通用的，不用重复写。

标签 PDF 里出现的每一个标签，核对结果只有三种：
    - 发货计划表里查不到这个标签「仓库含站点代号 + 未发货」的记录（比如已经发过了/发去别的
      仓库了/根本不在计划里）：跳过，不拆出来（这些箱子这次发货用不上）。
    - 查到了，但匹配到的待发行「工厂」不一致：跳过，记进报告，交给人工核对（不猜哪个是对的）。
    - 查到了且工厂一致：拆出来，箱数按查到的这几行「箱数」求和命名。
一个箱子如果混装了不止一个标签（比如 CA1 的 Packing List 有好几行），这个箱子的页面会出现
在每一个匹配上的标签的输出文件里——物理上没法把一箱货拆成两份，只能让涉及到的每个标签的
文件里都留一份完整的箱子页面，具体怎么处理这种箱子交给人工决定。
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import fitz

from .box import Box, LabelPdfStructureError
from .shipping_plan import load_pending_groups

ParseBoxesFn = Callable[[fitz.Document], list[Box]]


@dataclass
class SplitOutput:
    label: str
    factory: str
    total_boxes: float
    boxes_exact: bool
    box_count: int
    output_path: Path


@dataclass
class SplitReport:
    outputs: list[SplitOutput] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def derive_warehouse_code(label_pdf_path: str | Path) -> str:
    stem = Path(label_pdf_path).stem
    return stem.split("-", 1)[0].strip()


def run(label_pdf_path: str | Path, shipping_plan_path: str | Path, parse_boxes: ParseBoxesFn) -> SplitReport:
    label_pdf_path = Path(label_pdf_path)
    warehouse_code = derive_warehouse_code(label_pdf_path)

    doc = fitz.open(label_pdf_path)
    try:
        try:
            boxes = parse_boxes(doc)
        except LabelPdfStructureError as exc:
            raise LabelPdfStructureError(f"{label_pdf_path.name}：{exc}") from exc

        # 标签 -> box_no -> 这个箱子占的页码；用 box_no 去重，避免同一个箱子因为混装、
        # 明细里出现了不止一行同款标签而被重复收录
        pages_by_label: dict[str, dict[str, tuple[int, ...]]] = defaultdict(dict)
        for box in boxes:
            for item in box.items:
                pages_by_label[item.sku][box.box_no] = box.pages

        pdf_labels = set(pages_by_label)
        pending = load_pending_groups(shipping_plan_path, wanted_labels=pdf_labels, warehouse_code=warehouse_code)

        report = SplitReport()
        for label in sorted(pdf_labels - set(pending.groups) - set(pending.conflicts)):
            report.notes.append(
                f"「{label}」：标签 PDF 里有这个标签的箱子，但发货计划表里查不到仓库含 {warehouse_code} + 未发货的记录，跳过"
            )
        for label, factories in sorted(pending.conflicts.items()):
            report.notes.append(f"「{label}」：发货计划表里匹配到的待发行「工厂」不一致（{'、'.join(factories)}），跳过，需要人工核对")

        matched_labels = sorted(pending.groups)

        for label in matched_labels:
            group = pending.groups[label]
            # 按页码顺序还原箱子在原 PDF 里的先后顺序——不能按 box_no 字符串排序，box_no
            # 结尾的箱子序号是"1、2、…、10、11"这种数字，字符串排序会把 10 排到 2 前面
            ordered_boxes = sorted(pages_by_label[label].values(), key=lambda pages: pages[0])

            out_doc = fitz.open()
            for pages in ordered_boxes:
                for p in pages:
                    out_doc.insert_pdf(doc, from_page=p, to_page=p)

            out_dir = label_pdf_path.parent / _sanitize(group.factory) / _sanitize(warehouse_code)
            out_dir.mkdir(parents=True, exist_ok=True)

            boxes_text = str(group.total_boxes) if group.boxes_exact else f"{group.total_boxes:.1f}"
            out_name = f"{_sanitize(label)} {boxes_text}箱.pdf"
            out_path = out_dir / out_name
            out_doc.save(out_path)
            out_doc.close()

            report.outputs.append(
                SplitOutput(
                    label=label,
                    factory=group.factory,
                    total_boxes=group.total_boxes,
                    boxes_exact=group.boxes_exact,
                    box_count=len(ordered_boxes),
                    output_path=out_path,
                )
            )

        return report
    finally:
        doc.close()


def _sanitize(name: str) -> str:
    cleaned = "".join(c if c.isalnum() or c in "-_." else "_" for c in name)
    return cleaned or "unknown"
