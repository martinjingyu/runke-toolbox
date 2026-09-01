"""处理逻辑：对着真实的箱唛 PDF（"FBA标原文档" vs 已经人工处理过的"FBA标删除版"）核对过，
每页的"目的地"/"发货地"结构是固定的：

    目的地：                       发货地：
    FBA: <发货人名字>              <发货人名字>
    <FC 代码，比如 IND9>            <地址第二行>
    <街道地址>                     <地址第三行>
    <城市, 州 邮编>                 <国家（固定是"中国"）>
    <国家，比如"美国"/"加拿大">

规则（跟业务人员核对过实际例子，见同目录下"FBA标删除版"）：
    - 目的地：不管哪个国家，都把发货人名字删掉，FC 代码顶上去跟"FBA:"接在一起，
      后面几行（街道/城市州邮编/国家）各自上移一行。
    - 发货地：目的地国家是"美国"的话，删掉发货人名字这一行，后面几行上移一行，
      "发货地："这个标签本身保留；国家是"加拿大"的话，连"发货地："这个标签带整块内容
      全部删掉，不留痕迹。
    - 其他国家（不是美国也不是加拿大）：业务人员没说要怎么处理，这一页不动，报告里
      标出来"未知目的地国家"，交给人工看。
    - 目的地/发货地的行数结构跟预期不一样（比如不是 5 行/4 行）：这一页也不动，
      标"结构不符合预期"，不瞎猜着改。

跟参考例子（"FBA标删除版"）比，有一处简化：参考例子里发货地删除第一行之后，剩下的地址
文字整体重新排版折行（"Guangdong - dongguanshi -" 断成两行、跟下一行文字挤在一起），
这里没有照抄那个折行算法，而是让剩下几行各自成行、往上移——效果是一样的（发货人名字
不见了），版式更整齐，只是没有做到跟参考例子逐字节一致。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import fitz

_LINE_HEIGHT = 8.0
_FONT_SIZE = 8.0
_BASELINE_OFFSET = 6.5  # insert_text 的落点是文字基线，这个偏移量是从行框顶部换算成基线用的
_KNOWN_COUNTRIES = {"美国", "加拿大"}


@dataclass
class PageResult:
    file_name: str
    page_index: int
    status: str  # "美国" / "加拿大" / "未知目的地国家：xxx" / "结构不符合预期"
    modified: bool


@dataclass
class RunReport:
    results: list[PageResult] = field(default_factory=list)
    output_paths: list[Path] = field(default_factory=list)

    @property
    def modified_count(self) -> int:
        return sum(1 for r in self.results if r.modified)

    @property
    def skipped_results(self) -> list[PageResult]:
        return [r for r in self.results if not r.modified]


def find_pdfs(directory: str | Path) -> list[Path]:
    return sorted(Path(directory).glob("*.pdf"))


def _is_ascii(text: str) -> bool:
    return all(ord(c) < 128 for c in text)


def _draw_line(page: fitz.Page, x: float, top_y: float, text: str) -> None:
    font = "Helvetica" if _is_ascii(text) else "china-s"
    page.insert_text((x, top_y + _BASELINE_OFFSET), text, fontname=font, fontsize=_FONT_SIZE, color=(0, 0, 0))


def _get_lines(page: fitz.Page) -> list[dict]:
    lines = []
    for block in page.get_text("dict")["blocks"]:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            text = "".join(span["text"] for span in line["spans"])
            if text.strip():
                lines.append({"text": text, "bbox": line["bbox"]})
    return lines


def _lines_below(lines: list[dict], label: dict, max_count: int) -> list[dict]:
    same_column = [
        l for l in lines if abs(l["bbox"][0] - label["bbox"][0]) < 2 and l["bbox"][1] > label["bbox"][1]
    ]
    same_column.sort(key=lambda l: l["bbox"][1])
    return same_column[:max_count]


def redact_page(page: fitz.Page, file_name: str, page_index: int) -> PageResult:
    lines = _get_lines(page)

    dest_label = next((l for l in lines if l["text"].strip() == "目的地："), None)
    origin_label = next((l for l in lines if l["text"].strip() == "发货地："), None)
    if dest_label is None or origin_label is None:
        return PageResult(file_name, page_index, "找不到「目的地」/「发货地」标签", modified=False)

    dest_lines = _lines_below(lines, dest_label, 5)
    origin_lines = _lines_below(lines, origin_label, 4)
    if len(dest_lines) != 5 or len(origin_lines) != 4:
        return PageResult(file_name, page_index, "结构不符合预期（行数不对）", modified=False)

    # 光看行数不够——万一某页凑巧也是 5 行/4 行，但第一行其实不是发货人名字（比如地址本身
    # 写了两行），光按位置删就会删错东西。这里加一道结构层面的校验：目的地第一行必须是
    # "FBA:"开头——这是箱唛模板固定的格式，跟发货人具体叫什么名字无关，不管名字怎么变都稳定。
    # 不校验发货地第一行是不是跟目的地那边的名字文字完全一致——不同发货人名字写法本来就会不一样，
    # 强行要求两边一字不差反而会把本该处理的正常页面也拦下来。
    dest_line0 = " ".join(dest_lines[0]["text"].split())
    if not dest_line0.startswith("FBA:"):
        return PageResult(file_name, page_index, "目的地第一行不是「FBA:」开头，结构跟预期不符", modified=False)

    country = dest_lines[4]["text"].strip()
    if country not in _KNOWN_COUNTRIES:
        return PageResult(file_name, page_index, f"未知目的地国家：{country}", modified=False)

    # 擦除框的上边界不能从"内容第一行的顶部"算——实测第一行（发货人名字那行，字体被缩小
    # 到能塞下的程度）跟上面"目的地："/"发货地："标签的 bbox 在 Y 方向有一点重叠，擦除框只要
    # 沾到标签的 bbox 一点点，PyMuPDF 就会把整个标签文字也一起删掉。改成从标签的*底部*算起，
    # 不去碰标签自己的范围。右边界也不能写死一个很宽的数，会越界擦到右边那一栏的内容
    # （目的地在左栏、发货地在右栏，是并排的），改成按这个块里实际最宽的一行来算。
    x0 = dest_label["bbox"][0]
    dest_top = dest_lines[0]["bbox"][1]
    dest_right = max(l["bbox"][2] for l in dest_lines) + 2
    dest_rect = fitz.Rect(x0 - 1, dest_label["bbox"][3], dest_right, dest_lines[4]["bbox"][3] + 1)
    page.add_redact_annot(dest_rect, fill=(1, 1, 1))

    ox0 = origin_label["bbox"][0]
    origin_right = max(l["bbox"][2] for l in origin_lines) + 2
    if country == "加拿大":
        # 连"发货地："这个标签本身也删掉，整块清空，不重画任何内容——这里就要盖住标签自己
        # 的范围了，上边界用标签顶部
        origin_rect = fitz.Rect(ox0 - 1, origin_label["bbox"][1] - 1, origin_right, origin_lines[3]["bbox"][3] + 1)
    else:
        origin_rect = fitz.Rect(ox0 - 1, origin_label["bbox"][3], origin_right, origin_lines[3]["bbox"][3] + 1)
    page.add_redact_annot(origin_rect, fill=(1, 1, 1))

    page.apply_redactions()

    fc_code = dest_lines[1]["text"].strip()
    new_dest = [f"FBA: {fc_code}"] + [l["text"] for l in dest_lines[2:5]]
    for i, text in enumerate(new_dest):
        _draw_line(page, x0, dest_top + i * _LINE_HEIGHT, text)

    if country != "加拿大":
        origin_top = origin_lines[0]["bbox"][1]
        new_origin = [l["text"] for l in origin_lines[1:4]]
        for i, text in enumerate(new_origin):
            _draw_line(page, ox0, origin_top + i * _LINE_HEIGHT, text)

    return PageResult(file_name, page_index, country, modified=True)


def run(
    input_dir: str | Path,
    output_dir: str | Path,
    progress_callback: Callable[[int, int], None] | None = None,
) -> RunReport:
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_paths = find_pdfs(input_dir)
    total_pages = 0
    page_counts = []
    for path in pdf_paths:
        doc = fitz.open(path)
        page_counts.append(doc.page_count)
        total_pages += doc.page_count
        doc.close()

    report = RunReport()
    done_pages = 0

    for path, page_count in zip(pdf_paths, page_counts):
        doc = fitz.open(path)
        for i in range(page_count):
            result = redact_page(doc[i], path.name, i)
            report.results.append(result)
            done_pages += 1
            if progress_callback is not None:
                progress_callback(done_pages, total_pages)

        out_path = output_dir / path.name
        doc.save(out_path)
        doc.close()
        report.output_paths.append(out_path)

    return report
