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

重画的文字要跟原文档保持一样的字体/字号，不能瞎猜：实测这些箱唛只用两种字体——纯英文的
"FBA:"标签、单号这些是 Helvetica，其余（包括"YYC4"这种看着是纯 ASCII 的 FC 代码）都是
"STSong-Light"，且每一行的实际字号并不都是 8（比如地址太长会自动缩小），所以每行重画时
都直接沿用它自己在原文档里的字体名字/字号（从 get_text("dict") 的 span 里读出来），不再
按"是不是纯 ASCII"去猜该用哪种字体、也不再固定写死 8pt。

"STSong-Light"这个字体本身在原 PDF 里是没嵌入的（标准 14 种 CJK 字体之一，画的时候只写了
名字，指望阅读器自己有这个字体去替换显示）——PyMuPDF 自带的"china-s"这个替代名字实际指向
的是内置的"Droid Sans Fallback"（无衬线），跟"STSong-Light"（宋体，衬线）长得不一样，这就是
之前重画出来的字看着不对的根源。这里改成直接嵌入本机 Windows 自带的 STSong/新宋体字体文件
（C:／Windows／Fonts／STSONG.TTF，装了中文语言包的 Windows 机器上一般都有），跟原文档其它
没动过的文字用的是同一款字体，肉眼看不出区别；实在找不到这个字体文件的机器上，退回"china-s"，
好歹能画出字，不会因为缺字体直接失败。

加拿大目的地的 PDF 还有一层额外处理：一个文件里经常汇总了好几个厂商的货，脱敏之后还要
按每页的 SKU 查出厂商、把同一个厂商的页面拆到单独文件里，见 ca_split.py。
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import fitz

from .ca_split import VendorLookup, split_ca_pdf

_LINE_HEIGHT = 8.0
_BASELINE_RATIO = 6.5 / 8.0  # insert_text 的落点是文字基线，这个比例是从行框顶部换算成基线用的
# （原来是按固定 8pt 字号反推出的常数 6.5，现在字号跟着原文档每行实际字号变，换算成比例）
_KNOWN_COUNTRIES = {"美国", "加拿大"}

_CJK_FONT_NAME = "rk-fba-cjk"
_CJK_FONT_CANDIDATES = [
    r"C:\Windows\Fonts\STSONG.TTF",  # 跟原文档里没嵌入的"STSong-Light"是同一款字体，最匹配
    r"C:\Windows\Fonts\simsun.ttc",  # 找不到就退而求其次，用新宋体（视觉上很接近）
]


def _find_cjk_font_path() -> str | None:
    for path in _CJK_FONT_CANDIDATES:
        if os.path.exists(path):
            return path
    return None


_CJK_FONT_PATH = _find_cjk_font_path()  # 进程启动时找一次，不用每画一行字就查一次文件系统
_CJK_FONT = fitz.Font(fontfile=_CJK_FONT_PATH) if _CJK_FONT_PATH is not None else None
# 量文字宽度用的——fitz.get_text_length() 只认 Base14/CJK 保留名字这些"标准字体"，认不出自定义
# fontfile，得用 fitz.Font 对象自己的 text_length()，所以这里把字体对象也缓存一份


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
    notes: list[str] = field(default_factory=list)

    @property
    def modified_count(self) -> int:
        return sum(1 for r in self.results if r.modified)

    @property
    def skipped_results(self) -> list[PageResult]:
        return [r for r in self.results if not r.modified]


def find_pdfs(directory: str | Path) -> list[Path]:
    return sorted(Path(directory).glob("*.pdf"))


def _is_latin_font(source_font: str) -> bool:
    return "helvetica" in source_font.lower()


def _draw_run(page: fitz.Page, x: float, baseline_y: float, text: str, source_font: str, size: float) -> float:
    """画一段文字，返回这段文字的宽度——给"一行里要接着拼另一种字体的文字"这种场景用，
    知道前一段画完到哪儿了，后一段才能紧接着从那个位置开始画。"""
    if _is_latin_font(source_font):
        page.insert_text((x, baseline_y), text, fontname="Helvetica", fontsize=size, color=(0, 0, 0))
        return fitz.get_text_length(text, fontname="Helvetica", fontsize=size)
    if _CJK_FONT_PATH is not None:
        page.insert_text(
            (x, baseline_y), text, fontname=_CJK_FONT_NAME, fontfile=_CJK_FONT_PATH, fontsize=size, color=(0, 0, 0)
        )
        return _CJK_FONT.text_length(text, fontsize=size)
    page.insert_text((x, baseline_y), text, fontname="china-s", fontsize=size, color=(0, 0, 0))
    return fitz.get_text_length(text, fontname="china-s", fontsize=size)


def _draw_line(page: fitz.Page, x: float, top_y: float, text: str, source_font: str, size: float) -> None:
    """按 source_font 判断这行文字在原文档里用的是拉丁字体还是中文字体，画的时候原样沿用，
    字号也用原文档这一行自己的字号（不同行字号本来就不一样，别再统一写死 8pt）。"""
    baseline_y = top_y + size * _BASELINE_RATIO
    _draw_run(page, x, baseline_y, text, source_font, size)


def _draw_dest_first_line(page: fitz.Page, x: float, top_y: float, fc_code: str, label_line: dict, code_line: dict) -> None:
    """目的地第一行是"FBA: "接 FC 代码拼出来的，这两段在原文档里经常是不同字体——"FBA: "
    这个前缀固定是 Helvetica，FC 代码沿用它自己那一行原来的字体（很多时候是宋体）。整行只套
    一种字体会看着不对：宋体的拉丁字形笔画比 Helvetica 粗，"FBA: "如果套成宋体，看着就会比
    原文档粗一圈、像是被加粗了。所以这里分两段画，各自沿用自己的原始字体。

    字号统一用 FC 代码那一行的字号，不用标签那一行的——标签那一行原来的字号是因为要塞下很长
    的发货人名字才被压小的，新内容只有"FBA: <FC代码>"，长度跟原来对不上，沿用会显得偏小。
    """
    size = code_line["size"]
    baseline_y = top_y + size * _BASELINE_RATIO
    prefix_width = _draw_run(page, x, baseline_y, "FBA: ", label_line["font"], size)
    _draw_run(page, x + prefix_width, baseline_y, fc_code, code_line["font"], size)


_XOBJECT_DO_RE = re.compile(rb"/[A-Za-z0-9#_.\-]+ Do\b")


def _reset_text_render_mode_before_xobjects(page: fitz.Page) -> None:
    """apply_redactions() 内部会重整页面的内容流（clean_contents），实测这一步有时会把
    箱唛模板里"先切到描边模式画一段白色隐藏水印字、再切回填充模式"这一小段的"切回填充模式"
    重置操作漏掉——这个没被重置的描边状态会一路泄漏到后面用 Do 调用的内嵌图形对象里（比如
    条形码下面那行编号，是画在一个 Form XObject 里的），让那行字看着像被加粗了描边一样。这跟
    我们具体擦了哪块内容、画了什么字都没关系，纯粹是 apply_redactions 自己清理内容流时的副
    作用（对着完全没加任何擦除框的页面单独调用 clean_contents 也能复现）。

    这里的做法：每次 apply_redactions 之后，在页面内容流里每一处调用 Form/Image XObject 的
    "Do" 前面强制补一个"0 Tr"（纯填充、不描边）。不管前面泄漏了什么文字渲染模式，调用内嵌
    图形对象之前先归零——被调用的 XObject 自己内部要画字的话，本来就该自己设置字体/渲染模式
    （这份箱唛模板确实是这样做的），不会指望从外面继承来的状态，所以补这个重置不会影响正常
    内容，只会切断本不该泄漏过去的状态。
    """
    xrefs = page.get_contents()
    if not xrefs:
        return
    content = page.read_contents()
    patched = _XOBJECT_DO_RE.sub(lambda m: b"0 Tr " + m.group(0), content)
    if patched == content:
        return
    page.parent.update_stream(xrefs[0], patched)
    for xref in xrefs[1:]:
        page.parent.update_stream(xref, b"")


def _get_lines(page: fitz.Page) -> list[dict]:
    lines = []
    for block in page.get_text("dict")["blocks"]:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            text = "".join(span["text"] for span in line["spans"])
            if text.strip():
                first_span = line["spans"][0]
                lines.append(
                    {
                        "text": text,
                        "bbox": line["bbox"],
                        "font": first_span["font"],
                        "size": first_span["size"],
                    }
                )
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
    _reset_text_render_mode_before_xobjects(page)

    fc_code = dest_lines[1]["text"].strip()
    _draw_dest_first_line(page, x0, dest_top, fc_code, dest_lines[0], dest_lines[1])
    for i, l in enumerate(dest_lines[2:5]):
        _draw_line(page, x0, dest_top + (i + 1) * _LINE_HEIGHT, l["text"], l["font"], l["size"])

    if country != "加拿大":
        origin_top = origin_lines[0]["bbox"][1]
        for i, l in enumerate(origin_lines[1:4]):
            _draw_line(page, ox0, origin_top + i * _LINE_HEIGHT, l["text"], l["font"], l["size"])

    return PageResult(file_name, page_index, country, modified=True)


def run(
    input_dir: str | Path,
    output_dir: str | Path,
    progress_callback: Callable[[int, int], None] | None = None,
    vendor_lookup: VendorLookup | None = None,
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
        doc_results = []
        for i in range(page_count):
            result = redact_page(doc[i], path.name, i)
            doc_results.append(result)
            report.results.append(result)
            done_pages += 1
            if progress_callback is not None:
                progress_callback(done_pages, total_pages)

        is_ca = any(r.status == "加拿大" for r in doc_results)

        if is_ca and vendor_lookup is not None:
            split_result = split_ca_pdf(doc, path.name, output_dir, vendor_lookup)
            if split_result.note is None:
                report.output_paths.extend(split_result.output_paths)
                doc.close()
                continue
            report.notes.append(split_result.note)
            # 拆分失败（找不到 SKU/查不到厂商/文件名不符合规则），按原来的合并版输出，不丢文件

        if is_ca and vendor_lookup is None:
            report.notes.append(f"{path.name}：发往加拿大，汇总了多个厂商的货，但没有提供产品信息表，未按厂商拆分")

        out_path = output_dir / path.name
        doc.save(out_path)
        doc.close()
        report.output_paths.append(out_path)

    return report
