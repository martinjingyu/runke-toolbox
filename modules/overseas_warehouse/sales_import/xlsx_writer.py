"""往 xlsx 文件里改几个格子的值，除了这几个格子，其它字节原样不动。

为什么不能用 openpyxl 存盘：这份汇总表每一行的"图片"列用的是 WPS 专有的"单元格图片"
功能（`=_xlfn.DISPIMG(...)`），图片数据实际存在 `xl/cellimages.xml` + `xl/media/*`
这两个 openpyxl 完全不认识的部件里。实测过——哪怕一个格子都不改，只要 openpyxl
`load_workbook()` 再 `save()`，这些部件会被整个丢弃（106MB 的文件存完只剩 232KB，
所有商品图直接没了）。所以只能不经过 openpyxl 的对象模型，直接在 xlsx 这个 zip 包里
定位到要改的 sheet XML，用字符串替换改掉具体的 `<c>` 格子节点，其它 zip 里的部件
（图片、样式、其它 sheet、公式……）整个原样拷贝过去。

这里只处理"写一个纯数字值"这一种场景（销量格子），不处理公式格子、不处理字符串格子
（不需要碰 sharedStrings.xml）。

**改完之后必须强制整个工作簿在下次打开时重新算一遍公式**——这个坑是拿真实文件验证出来
才发现的：汇总表里"该行 Single SKU 销售总量"（比如 M4=SUM(N4:AR4)，把这一行31天的
销量加总）、"该平台 All SKU 销售总量"（第2行，比如=SUM(M4:M14)，把这个平台所有 SKU
的行汇总再加总）这些格子都是公式，我们这里只改了公式依赖的"某一天"那个格子的值，
公式本身没碰、但它缓存在 XML 里的 `<v>` 结果没有跟着重新算——Excel/WPS 是不是会自动
重算，取决于这份工作簿 `xl/workbook.xml` 的 `<calcPr>` 设置里有没有开
`fullCalcOnLoad`，实测这份表原本没开，导致这些汇总格子在业务人员打开表之后还是显示
改之前的旧值。修法不是自己去重新算这些公式（公式种类多、还有跨 sheet 的
VLOOKUP，自己算太容易算错），而是把 `fullCalcOnLoad` 打开，让打开这份表的 Excel/WPS
自己去重算全表——这是 OOXML 标准里就有的机制，不是我们发明的。
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass

import openpyxl.utils

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_CELL_REF_RE = re.compile(r'<c r="([A-Z]+)\d+"')


class XlsxPatchError(Exception):
    pass


@dataclass(frozen=True)
class CellUpdate:
    row: int
    col: int
    value: float


def _sheet_xml_path(zin: zipfile.ZipFile, sheet_name: str) -> str:
    import xml.etree.ElementTree as ET

    wb_root = ET.fromstring(zin.read("xl/workbook.xml"))
    sheets_el = wb_root.find(f"{{{_NS_MAIN}}}sheets")
    rid = None
    for s in sheets_el:
        if s.get("name") == sheet_name:
            rid = s.get(f"{{{_NS_R}}}id")
            break
    if rid is None:
        raise XlsxPatchError(f"工作簿里找不到叫 {sheet_name!r} 的 sheet")

    rels_root = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_root:
        if rel.get("Id") == rid:
            target = rel.get("Target").lstrip("/")
            return target if target.startswith("xl/") else "xl/" + target
    raise XlsxPatchError(f"workbook.xml.rels 里找不到 {sheet_name!r} 对应的关系 {rid!r}")


def _col_letters_to_index(letters: str) -> int:
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _replace_existing_cell(row_xml: str, col_ref: str, value: float) -> str | None:
    """命中已有的 <c> 节点（自闭合或者带 <v>）就地把值换掉，换不到返回 None。

    注意：一定要先试自闭合的 pattern，再试带值的 pattern——一个不区分 "/>" 和 ">"
    的宽松正则会把自闭合格子也当成"带值"格子去匹配，然后它非贪婪的 (.*?) 会一路
    吞掉后面所有自闭合的兄弟格子，直到撞上后面其它格子真正的 </c> 才停，等于把
    中间一串格子全部删掉。这个坑是拿真实文件测出来的，不是猜的。
    """
    self_closed_pat = re.compile(rf'<c r="{col_ref}"([^>]*?)/>')
    m = self_closed_pat.search(row_xml)
    if m:
        attrs = re.sub(r'\st="[^"]*"', "", m.group(1))
        new_cell = f'<c r="{col_ref}"{attrs}><v>{value}</v></c>'
        return row_xml[: m.start()] + new_cell + row_xml[m.end():]

    with_value_pat = re.compile(rf'<c r="{col_ref}"([^>]*?)>(.*?)</c>', re.DOTALL)
    m = with_value_pat.search(row_xml)
    if m:
        attrs = re.sub(r'\st="[^"]*"', "", m.group(1))
        new_cell = f'<c r="{col_ref}"{attrs}><v>{value}</v></c>'
        return row_xml[: m.start()] + new_cell + row_xml[m.end():]

    return None


def _insert_new_cell(row_xml: str, col_ref: str, col_index: int, value: float) -> str:
    """这一行里压根没有这一列的 <c> 节点（比如这个格子从没被设过任何格式）——
    按列号顺序插入一个新的、没有样式的 <c> 节点。"""
    new_cell = f'<c r="{col_ref}"><v>{value}</v></c>'
    for m in _CELL_REF_RE.finditer(row_xml):
        if _col_letters_to_index(m.group(1)) > col_index:
            return row_xml[: m.start()] + new_cell + row_xml[m.start():]
    end_tag = row_xml.rfind("</row>")
    if end_tag == -1:
        raise XlsxPatchError("row xml 里找不到 </row>，格式不对，不敢插入")
    return row_xml[:end_tag] + new_cell + row_xml[end_tag:]


def _patch_row(row_xml: str, col_updates: dict[int, float]) -> str:
    for col, value in col_updates.items():
        col_ref = f"{openpyxl.utils.get_column_letter(col)}{_row_num_of(row_xml)}"
        replaced = _replace_existing_cell(row_xml, col_ref, value)
        if replaced is not None:
            row_xml = replaced
        else:
            row_xml = _insert_new_cell(row_xml, col_ref, col, value)
    return row_xml


def _row_num_of(row_xml: str) -> str:
    m = re.match(r'<row r="(\d+)"', row_xml)
    if not m:
        raise XlsxPatchError("row xml 开头解析不出行号")
    return m.group(1)


_CALC_PR_RE = re.compile(r"<calcPr([^>]*)/>")
_FULL_CALC_ATTR_RE = re.compile(r'\s*fullCalcOnLoad="[^"]*"')


def _force_full_calc_on_load(workbook_xml_bytes: bytes) -> bytes:
    """给 xl/workbook.xml 的 <calcPr> 标签打开 fullCalcOnLoad="1"——告诉打开这份表的
    Excel/WPS："这份文件里有公式的缓存值可能不准了，打开的时候整表重新算一遍"。
    没有 calcPr 标签（少见，但 OOXML 规范里这个标签本来就是可选的）就自己加一个。
    """
    text = workbook_xml_bytes.decode("utf-8")
    m = _CALC_PR_RE.search(text)
    if m:
        attrs = _FULL_CALC_ATTR_RE.sub("", m.group(1))
        new_tag = f'<calcPr{attrs} fullCalcOnLoad="1"/>'
        text = text[: m.start()] + new_tag + text[m.end():]
    else:
        insert_at = text.rfind("</workbook>")
        if insert_at == -1:
            raise XlsxPatchError("workbook.xml 里找不到 </workbook>，格式不对")
        text = text[:insert_at] + '<calcPr fullCalcOnLoad="1"/>' + text[insert_at:]
    return text.encode("utf-8")


def _patch_sheet_xml(sheet_xml_bytes: bytes, updates: dict[int, dict[int, float]]) -> bytes:
    text = sheet_xml_bytes.decode("utf-8")
    for row_num, col_updates in updates.items():
        row_pat = re.compile(rf'<row r="{row_num}"[^>]*>.*?</row>', re.DOTALL)
        m = row_pat.search(text)
        if not m:
            raise XlsxPatchError(f"sheet xml 里找不到第 {row_num} 行——这一行理应已经存在真实数据，属于内部错误")
        patched_row = _patch_row(m.group(0), col_updates)
        text = text[: m.start()] + patched_row + text[m.end():]
    return text.encode("utf-8")


def apply_cell_updates(src_path: str, dest_path: str, updates: dict[str, dict[tuple[int, int], float]]) -> None:
    """把 updates 里指定的格子值写进 dest_path（src_path 复制过去的副本），
    除了被改的那几个 sheet 的 XML，zip 里其它所有部件原样拷贝，一个字节都不碰。

    updates: {sheet_name: {(row, col): new_value}}
    """
    with zipfile.ZipFile(src_path, "r") as zin:
        patched: dict[str, bytes] = {}
        for sheet_name, cell_updates in updates.items():
            if not cell_updates:
                continue
            target_path = _sheet_xml_path(zin, sheet_name)
            by_row: dict[int, dict[int, float]] = {}
            for (row, col), value in cell_updates.items():
                by_row.setdefault(row, {})[col] = value
            sheet_bytes = patched.get(target_path) or zin.read(target_path)
            patched[target_path] = _patch_sheet_xml(sheet_bytes, by_row)

        if patched:
            # 只要真的改了任何格子，就强制下次打开整表重算——不然改了销量格子，
            # 依赖它的"该行汇总"/"该平台汇总"这些公式格子会显示旧的缓存值。
            patched["xl/workbook.xml"] = _force_full_calc_on_load(zin.read("xl/workbook.xml"))

        with zipfile.ZipFile(dest_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = patched.get(item.filename, None)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
