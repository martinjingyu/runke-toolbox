"""解析物流部门每天从亚马逊后台爬到的 FBA 货件数据——一个文件夹下若干 CSV，一份文件对应
一个账号/站点，列都是同一套（货件名称、FBA货件编号、内部编号、目的地、MSKU、预计商品
数量、MSKU状态）。

"货件名称"里带着这批货originally打算发货的日期，形如 `GH-9.9-SBD1`（月.日，不带年份）；
已经被取消的货件，货件名称会变成"取消"或者"取消"+原名称，"MSKU状态"也会同步变成"已取消"。
这两条（日期对不上筛选日期 / 已取消）是两条独立的跳过判断，不是非此即彼——都要单独记录
原因，最后汇总成一份"跳过说明"，不能悄悄丢掉不告诉人。

同一批货物什么原因不一定，可能会被不同日期抓到的 CSV 重复收录（内部编号/FBA货件编号完全
一样）——这种情况只算一次，不然后面按 MSKU 汇总数量会重复计数，跟发货计划表对不上。
"""
from __future__ import annotations

import csv
import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path

REQUIRED_COLUMNS = ["货件名称", "FBA货件编号", "内部编号", "目的地", "MSKU", "预计商品数量", "MSKU状态"]

CANCELLED_STATUS = "已取消"

# 货件名称形如 "GH-9.9-SBD1"/"取消TZ-9.16-SBD1"/"GH-9.9--TEB9"（偶尔有双短横线）——不锚定
# 开头结尾，只找第一处"数字.数字"，能兼容各种前缀/后缀写法。
_DATE_RE = re.compile(r"(\d{1,2})\.(\d{1,2})")


class FbaSourceError(Exception):
    pass


@dataclass
class FbaEntry:
    msku: str
    destination: str
    fba_id: str
    internal_id: str
    quantity_pieces: int
    source_file: str
    source_row: int


@dataclass
class SkippedFbaRow:
    source_file: str
    source_row: int
    shipment_name: str
    msku: str
    reason: str


@dataclass
class FbaSourceResult:
    entries_by_msku: dict[str, list[FbaEntry]] = field(default_factory=dict)
    skipped: list[SkippedFbaRow] = field(default_factory=list)
    duplicates_removed: int = 0


def _parse_month_day(shipment_name: str) -> tuple[int, int] | None:
    match = _DATE_RE.search(shipment_name)
    if match is None:
        return None
    month, day = int(match.group(1)), int(match.group(2))
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    return month, day


def _detect_encoding(path: Path) -> str:
    """不同账号/设备爬出来的 CSV 编码不统一——实测同一个文件夹里，有的文件是带 BOM 的
    UTF-8，有的是 GBK（没有 BOM，Excel 另存 CSV 的默认编码），猜错了整份文件从第一行
    表头开始就会 UnicodeDecodeError。用 BOM 判断 UTF-8，判断不出来就退回 GBK——不能
    在整个文件夹层面固定用一种编码。
    """
    with open(path, "rb") as fh:
        head = fh.read(3)
    if head == b"\xef\xbb\xbf":
        return "utf-8-sig"
    return "gbk"


def _iter_csv_rows(path: Path):
    encoding = _detect_encoding(path)
    try:
        with open(path, "r", encoding=encoding, newline="") as fh:
            reader = csv.DictReader(fh)
            if reader.fieldnames is None:
                raise FbaSourceError(f"「{path.name}」是空文件，没有表头")
            missing = [c for c in REQUIRED_COLUMNS if c not in reader.fieldnames]
            if missing:
                raise FbaSourceError(f"「{path.name}」缺少必须的表头：{missing}")
            rows = list(enumerate(reader, start=2))  # 第 1 行是表头，数据从第 2 行开始
    except UnicodeDecodeError as exc:
        raise FbaSourceError(f"「{path.name}」按 {encoding} 编码读取失败：{exc}") from exc
    yield from rows


def parse_fba_folder(folder: Path, ship_date: dt.date) -> FbaSourceResult:
    """扫一遍文件夹下所有 *.csv，按文件名排序（保证多次运行顺序稳定），逐行判断要不要跳过，
    保留下来的行按 MSKU 分组、组内按遇到的先后顺序排列——后面"拉链式"拆分时，FBA 号谁先谁后
    没有业务要求（已经跟用户确认过），这里只要顺序稳定、可复现就行。
    """
    folder = Path(folder)
    csv_paths = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".csv")
    if not csv_paths:
        raise FbaSourceError(f"「{folder}」下没有找到任何 CSV 文件")

    result = FbaSourceResult()
    seen_keys: set[tuple] = set()

    for path in csv_paths:
        for row_idx, row in _iter_csv_rows(path):
            shipment_name = (row.get("货件名称") or "").strip()
            msku = (row.get("MSKU") or "").strip()
            status = (row.get("MSKU状态") or "").strip()
            destination = (row.get("目的地") or "").strip()
            fba_id = (row.get("FBA货件编号") or "").strip()
            internal_id = (row.get("内部编号") or "").strip()
            qty_raw = (row.get("预计商品数量") or "").strip()

            def skip(reason: str) -> None:
                result.skipped.append(
                    SkippedFbaRow(
                        source_file=path.name,
                        source_row=row_idx,
                        shipment_name=shipment_name,
                        msku=msku,
                        reason=reason,
                    )
                )

            reasons = []
            if status == CANCELLED_STATUS:
                reasons.append("MSKU状态=已取消")

            parsed = _parse_month_day(shipment_name)
            if parsed is None:
                reasons.append(f"货件名称「{shipment_name}」里解析不出日期")
            elif parsed != (ship_date.month, ship_date.day):
                reasons.append(
                    f"货件名称日期 {parsed[0]}.{parsed[1]} 跟筛选日期 "
                    f"{ship_date.month}.{ship_date.day} 不一致"
                )

            if reasons:
                skip("；".join(reasons))
                continue

            try:
                quantity = int(float(qty_raw))
            except ValueError:
                skip(f"预计商品数量「{qty_raw}」不是数字")
                continue

            if quantity == 0:
                # 正常情况下数量=0 的行都是已取消（上面已经拦掉了），这里走到说明状态没标
                # 已取消但数量确实是 0——没有货可分配，没必要往下参与拆分，跳过并单独说明，
                # 不能悄悄当成"跟已取消一样处理"混进去，人工应该知道有这么一条异常记录。
                skip("预计商品数量为 0，但 MSKU状态 不是已取消")
                continue

            dedup_key = (msku, destination, fba_id, internal_id, quantity)
            if dedup_key in seen_keys:
                result.duplicates_removed += 1
                continue
            seen_keys.add(dedup_key)

            result.entries_by_msku.setdefault(msku, []).append(
                FbaEntry(
                    msku=msku,
                    destination=destination,
                    fba_id=fba_id,
                    internal_id=internal_id,
                    quantity_pieces=quantity,
                    source_file=path.name,
                    source_row=row_idx,
                )
            )

    return result
