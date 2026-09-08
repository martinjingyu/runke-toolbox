"""加拿大目的地的箱唛 PDF，一个文件里汇总了好几个厂商的货（比如原始文件名
"GH-WJ-9.9-CA-YYC4 FBA19NVYQBMT.pdf"，GH、WJ 是两个厂商代号），需要按每页实际装的货
拆成一个厂商一个文件，不能整份原样发给随便哪个厂商。

拆分依据是每页箱唛里"Single SKU"下面那一行的编号（跟"数量"下面重复的那个编号是同一个
值，取"Single SKU"下面那个就够了）——拿这个编号去产品信息总表里查 RK-SKU/AMZ-SKU 对应的
"厂商"这一列（查法跟 shipment_plan_apply/product_lookup.py 一致：先按 RK-SKU 查，查不到
再按 AMZ-SKU 查），得到厂商代号，同一个厂商代号的页面合并进同一个输出文件。

原文件名里"厂商代号们"和"其余部分"的分界，用日期段（形如"9.9"这种"数字.数字"）来找——
这是箱唛命名里唯一稳定、不会跟厂商代号自己长得一样的标记。分界之前的部分（比如
"GH-WJ"）就是这一批 PDF 里出现的厂商范围，分界开始的部分（日期/国家/FC/空格/FBA单号/
扩展名）原样保留：GH-WJ-9.9-CA-YYC4 FBA19NVYQBMT.pdf 拆出来就是
GH-9.9-CA-YYC4 FBA19NVYQBMT.pdf 和 WJ-9.9-CA-YYC4 FBA19NVYQBMT.pdf。

文件名里的"厂商范围"不只是用来拼输出文件名，还是一道交叉校验：产品信息表里同一个 SKU
完全可能对应不止一个厂商（不同厂商各自生产同一个型号很正常），单看表查出来的可能是好几个
候选厂商，光凭表本身没法确定这一页具体是哪个厂商发的——这时候用文件名里已经列出的厂商范围
去交叉一下：候选厂商里正好只有一个落在这个范围内，就是它；候选厂商一个都不在这个范围内，
或者交叉完还剩不止一个，都说明有问题，不能瞎猜，报错交给人工核对（可能是产品信息表数据错了，
也可能是这批 PDF 混进了不该在这个范围里的货）。

任何一步查不清楚（页面上找不到 SKU、SKU 在表里一个厂商都查不到、查到的厂商都不在文件名列出
的范围内、交叉完还是有歧义、文件名里找不到日期段分界）都不硬拆——原样按合并版输出，在报告里
写清楚原因，交给人工处理，跟 redact.py 里"结构不符合预期就不动"是同一个原则。
"""
from __future__ import annotations

import re
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path

import fitz

from ..shipment_plan_apply.column_utils import column_index_map, find_header_row, require_columns

VENDOR_LOOKUP_REQUIRED_HEADERS = ["AMZ-SKU", "RK-SKU", "厂商"]

_VENDOR_DATE_SEGMENT_RE = re.compile(r"^\d+\.\d+$")


@dataclass
class VendorLookup:
    amz_to_vendors: dict[str, set[str]] = field(default_factory=dict)
    rk_to_vendors: dict[str, set[str]] = field(default_factory=dict)

    def resolve_candidates(self, sku: str) -> set[str]:
        """一个 SKU 在表里可能对应不止一个厂商（不同厂商各自生产同一型号），这里如实返回
        全部候选，交给调用方结合文件名里的厂商范围去交叉判断，不在这里挑一个了事。"""
        vendors = self.rk_to_vendors.get(sku)
        if vendors:
            return vendors
        return self.amz_to_vendors.get(sku, set())


def load_vendor_lookup(path: str | Path) -> VendorLookup:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    header_row = find_header_row(ws, VENDOR_LOOKUP_REQUIRED_HEADERS, context="产品信息表")
    cols = column_index_map(ws, header_row)
    idx = require_columns(cols, VENDOR_LOOKUP_REQUIRED_HEADERS, "产品信息表")
    amz_col = idx["AMZ-SKU"]
    rk_col = idx["RK-SKU"]
    vendor_col = idx["厂商"]

    amz_to_vendors: dict[str, set[str]] = {}
    rk_to_vendors: dict[str, set[str]] = {}

    for row in ws.iter_rows(min_row=header_row + 1):
        amz = row[amz_col - 1].value
        rk = row[rk_col - 1].value
        vendor = row[vendor_col - 1].value

        if vendor is None:
            continue
        vendor = str(vendor).strip()
        if not vendor:
            continue

        if amz is not None:
            amz = str(amz).strip()
            if amz:
                amz_to_vendors.setdefault(amz, set()).add(vendor)

        if rk is not None:
            rk = str(rk).strip()
            if rk:
                rk_to_vendors.setdefault(rk, set()).add(vendor)

    return VendorLookup(amz_to_vendors=amz_to_vendors, rk_to_vendors=rk_to_vendors)


def extract_sku(page: fitz.Page) -> str | None:
    lines = [line.strip() for line in page.get_text().splitlines() if line.strip()]
    for i, line in enumerate(lines):
        if line == "Single SKU" and i + 1 < len(lines):
            return lines[i + 1]
    return None


def split_vendor_prefix(file_name: str) -> tuple[list[str], str] | None:
    """按日期段（"数字.数字"）分界，把文件名拆成前面的厂商代号列表和后面原样保留的部分。
    找不到这样的分界就返回 None。"""
    parts = file_name.split("-")
    for i, part in enumerate(parts):
        if _VENDOR_DATE_SEGMENT_RE.match(part):
            vendors = parts[:i]
            rest = "-".join(parts[i:])
            if vendors and rest:
                return vendors, rest
            return None
    return None


def extract_warehouse_code(rest: str) -> str | None:
    """rest 是 split_vendor_prefix 拆出来的"从日期段开始的部分"，形如"9.9-CLT2 FBA19NW5LTHC
     NK.pdf"（美国）或"9.9-CA-YYC4 FBA19NVYQBMT.pdf"（加拿大，日期段后面多一段"CA"国家
    标记）。目的地仓库/FC 代号紧跟在日期段（美国）或"CA"标记（加拿大）后面，取空格之前的那
    一段——命名习惯里这几段都是用"-"连接、FC 代号后面用空格接单号/其它文字。取不出来（结构
    跟预期不符，比如日期段后面没东西了）返回 None，交给调用方决定怎么处理，不硬猜。
    """
    parts = rest.split("-")
    if len(parts) < 2:
        return None
    idx = 1
    if parts[idx].strip() == "CA":
        idx += 1
    if idx >= len(parts):
        return None
    tokens = parts[idx].split()
    return tokens[0] if tokens else None


def resolve_vendor_and_warehouse(file_name: str) -> tuple[str, str] | None:
    """从文件名（还没拆分过的原始文件名，或者已经拆成单厂商的文件名都行）里识别出唯一的
    厂商代号和目的地仓库/FC 代号——给"按厂商+仓库分文件夹"用。文件名里的厂商代号不止一个
    （比如"GH-WJ-..."这种还没按厂商拆开的加拿大汇总文件），或者压根找不到日期分段/仓库代号，
    都返回 None，不硬选一个，交给调用方决定怎么兜底（不分类，原样放在顶层目录）。
    """
    prefix = split_vendor_prefix(file_name)
    if prefix is None:
        return None
    vendors, rest = prefix
    if len(vendors) != 1:
        return None
    warehouse = extract_warehouse_code(rest)
    if warehouse is None:
        return None
    return vendors[0], warehouse


def extract_pages(src: fitz.Document, page_indices: list[int]) -> fitz.Document:
    new_doc = fitz.open()
    for idx in page_indices:
        new_doc.insert_pdf(src, from_page=idx, to_page=idx)
    return new_doc


@dataclass
class CaSplitResult:
    output_paths: list[Path]
    note: str | None  # 拆分失败时，写清楚原因；成功时是 None


def split_ca_pdf(doc: fitz.Document, file_name: str, output_dir: Path, vendor_lookup: VendorLookup) -> CaSplitResult:
    # 曾经试过让调用方传 resolve_output_dir(vendor, warehouse) -> Path，把每个厂商的拆分
    # 结果分别存到各自的"厂商/仓库"文件夹里——先改回最简单的"都存到同一个 output_dir"，
    # 想恢复按厂商/仓库分文件夹的话，参考 resolve_vendor_and_warehouse/extract_warehouse_code
    # 这两个函数（还留着，没删）。
    prefix = split_vendor_prefix(file_name)
    if prefix is None:
        return CaSplitResult([], f"{file_name}：文件名里找不到日期分段，无法确定厂商代号和其余部分的分界，未拆分")

    original_vendors, rest = prefix
    allowed_vendors = set(original_vendors)

    page_vendors: list[str] = []
    for i in range(doc.page_count):
        sku = extract_sku(doc[i])
        if sku is None:
            return CaSplitResult([], f"{file_name}：第 {i + 1} 页找不到「Single SKU」编号，未拆分")

        candidates = vendor_lookup.resolve_candidates(sku)
        if not candidates:
            return CaSplitResult([], f"{file_name}：第 {i + 1} 页的 SKU「{sku}」在产品信息表里查不到厂商，未拆分")

        # 产品信息表里查出来的候选厂商，跟文件名开头已经列出的厂商范围交叉一下：文件名
        # 是"这一批 PDF 里确实出现过哪些厂商"的已知事实，候选厂商如果都不在这个范围里，
        # 或者交叉完还剩不止一个，都说明查出来的结果跟已知事实对不上，不能瞎猜。
        matched = candidates & allowed_vendors
        if not matched:
            return CaSplitResult(
                [],
                f"{file_name}：第 {i + 1} 页的 SKU「{sku}」查到的厂商（{'、'.join(sorted(candidates))}）"
                f"都不在文件名列出的厂商范围（{'、'.join(sorted(allowed_vendors))}）内，未拆分",
            )
        if len(matched) > 1:
            return CaSplitResult(
                [],
                f"{file_name}：第 {i + 1} 页的 SKU「{sku}」同时匹配文件名里的多个厂商"
                f"（{'、'.join(sorted(matched))}），无法确定具体是哪一个，未拆分",
            )
        page_vendors.append(next(iter(matched)))

    groups: "OrderedDict[str, list[int]]" = OrderedDict()
    for i, vendor in enumerate(page_vendors):
        groups.setdefault(vendor, []).append(i)

    output_paths: list[Path] = []
    for vendor, indices in groups.items():
        new_doc = extract_pages(doc, indices)
        # 都存到同一个 output_dir，文件名带上厂商前缀区分，不然不同厂商拆出来的文件名会撞车
        # （拆分之前只有一份 rest，同一份原文件拆出的每个厂商版本 rest 都一样）
        out_path = output_dir / f"{vendor}-{rest}"
        new_doc.save(out_path)
        new_doc.close()
        output_paths.append(out_path)

    return CaSplitResult(output_paths, None)
