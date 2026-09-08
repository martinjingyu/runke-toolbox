"""拆分入口：先看标签 PDF（比如 CA1.pdf）里实际出现过哪些标签，再拿这份名单去发货计划表里
逐个核对——只查这些标签，不读发货计划表里其它几万行跟这份 PDF 无关的历史数据（那些标签根本
不在这份 PDF 里，读出来也没用，还会在报告里制造一堆无关的噪音）。核对到「仓库含 CA1、状态=
未发货」的记录（这个工具目前专门服务 CA1 这个站点，见 shipping_plan.py 里 WAREHOUSE_CODE
的说明；不要求发货时间=待定，已经安排了具体发货日期但还没实际发出去的也算），就把对应的箱子
页面（Inbound+Packing List 两页一组）从标签 PDF 里抽出来，按 厂商代号_标签_箱数合计.pdf
命名，存到标签 PDF 所在目录下新建的 output 子文件夹。

标签 PDF 里出现的每一个标签，核对结果只有三种：
    - 发货计划表里查不到这个标签「仓库含 CA1 + 未发货」的记录（比如已经发过了/发去别的
      仓库了/根本不在计划里）：跳过，不拆出来（这些箱子这次发货用不上）。
    - 查到了，但匹配到的待发行「工厂」不一致：跳过，记进报告，交给人工核对（不猜哪个是对的）。
    - 查到了且工厂一致：拆出来，箱数按查到的这几行「箱数」求和命名。
一个箱子如果混装了不止一个标签（Packing List 有好几行），这个箱子的两页会出现在每一个
匹配上的标签的输出文件里——物理上没法把一箱货拆成两份，只能让涉及到的每个标签的文件里都留
一份完整的箱子页面，具体怎么处理这种箱子交给人工决定。
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import fitz

from .label_pdf import LabelPdfStructureError, parse_label_pdf
from .shipping_plan import load_pending_groups

OUTPUT_DIR_NAME = "output"


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


def run(label_pdf_path: str | Path, shipping_plan_path: str | Path) -> SplitReport:
    label_pdf_path = Path(label_pdf_path)
    output_dir = label_pdf_path.parent / OUTPUT_DIR_NAME

    doc = fitz.open(label_pdf_path)
    try:
        try:
            boxes = parse_label_pdf(doc)
        except LabelPdfStructureError as exc:
            raise LabelPdfStructureError(f"{label_pdf_path.name}：{exc}") from exc

        # 标签 -> box_no -> (Inbound 页码, Packing List 页码)；用 box_no 去重，避免同一个箱子
        # 因为混装、Packing List 里出现了不止一行同款标签而被重复收录
        pages_by_label: dict[str, dict[str, tuple[int, int]]] = defaultdict(dict)
        for box in boxes:
            for item in box.items:
                pages_by_label[item.sku][box.box_no] = (box.inbound_page, box.packing_page)

        pdf_labels = set(pages_by_label)
        pending = load_pending_groups(shipping_plan_path, wanted_labels=pdf_labels)

        report = SplitReport()
        for label in sorted(pdf_labels - set(pending.groups) - set(pending.conflicts)):
            report.notes.append(f"「{label}」：标签 PDF 里有这个标签的箱子，但发货计划表里查不到仓库含 CA1 + 未发货的记录，跳过")
        for label, factories in sorted(pending.conflicts.items()):
            report.notes.append(f"「{label}」：发货计划表里匹配到的待发行「工厂」不一致（{'、'.join(factories)}），跳过，需要人工核对")

        matched_labels = sorted(pending.groups)
        if matched_labels:
            output_dir.mkdir(parents=True, exist_ok=True)

        for label in matched_labels:
            group = pending.groups[label]
            # 按页码顺序还原箱子在原 PDF 里的先后顺序——不能按 box_no 字符串排序，box_no
            # 结尾的箱子序号是"1、2、…、10、11"这种数字，字符串排序会把 10 排到 2 前面
            ordered_pairs = sorted(pages_by_label[label].values(), key=lambda pair: pair[0])

            out_doc = fitz.open()
            for inbound_idx, packing_idx in ordered_pairs:
                out_doc.insert_pdf(doc, from_page=inbound_idx, to_page=inbound_idx)
                out_doc.insert_pdf(doc, from_page=packing_idx, to_page=packing_idx)

            boxes_text = str(group.total_boxes) if group.boxes_exact else f"{group.total_boxes:.1f}"
            out_name = f"{_sanitize(group.factory)}_{_sanitize(label)}_{boxes_text}.pdf"
            out_path = output_dir / out_name
            out_doc.save(out_path)
            out_doc.close()

            report.outputs.append(
                SplitOutput(
                    label=label,
                    factory=group.factory,
                    total_boxes=group.total_boxes,
                    boxes_exact=group.boxes_exact,
                    box_count=len(ordered_pairs),
                    output_path=out_path,
                )
            )

        return report
    finally:
        doc.close()


def _sanitize(name: str) -> str:
    cleaned = "".join(c if c.isalnum() or c in "-_." else "_" for c in name)
    return cleaned or "unknown"
