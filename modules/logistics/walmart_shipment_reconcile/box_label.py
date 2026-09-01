"""解析箱唛 PDF。

每一页要么是一个箱子的标签（右下角有个 Data Matrix 二维码），要么是托盘汇总页
（"PALLET X OF Y"，只有一维条码，没有 GTIN/数量信息）——后者跳过，不参与核对。

二维码里的内容是按 GS1 风格的 AI（Application Identifier）连在一起的，实测格式是：
    400<SHIPMENT ID>01<14位GTIN>00<18位BOX ID>30<QUANTITY>
用真实标签核对过：SHIPMENT ID、GTIN、BOX ID、QUANTITY 都跟标签上印刷的文字一致，直接读条码
比识别图片文字准确得多，这几个字段都从条码来，不用 OCR；解不出来的页面按“非箱标签页”跳过。

条码里没有编码 SKU 文字本身（只有 GTIN），而这次能拿到的几份表都不含 GTIN，只能靠 SKU 文字
互相对应，所以 SKU 这一个字段没法走条码，要靠 ocr_single_sku() 读印刷体——每个 GTIN 只需要
OCR 一次代表页（同一个 GTIN 的箱子必然是同一个 SKU），不用每一页都读，准确率和速度都还好。

每一页“渲染成图片 + 解条码”都是独立的纯 CPU 计算（实测约 2.6 秒/页），互相不依赖，
所以按页数用多个进程并行处理——页数多的时候（比如几百页）能明显缩短总时间。
"""
from __future__ import annotations

import os
import platform
import re
import threading
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import fitz

if platform.system() == "Darwin":
    # Mac 上开发环境：pylibdmtx 靠 ctypes 的 find_library 找 libdmtx.dylib，
    # 默认搜索路径里不含 /opt/homebrew/lib（brew install libdmtx 装的位置），要手动加一下。
    # Windows 打包时用的是 pylibdmtx 自带的 dmtx.dll，不需要这一步。这段要在每个子进程里也执行到，
    # 所以放在模块顶层——子进程会重新 import 一遍这个模块，同样的逻辑会再跑一次。
    os.environ.setdefault("DYLD_LIBRARY_PATH", "/opt/homebrew/lib")

from pylibdmtx.pylibdmtx import decode as _dmtx_decode  # noqa: E402
from PIL import Image  # noqa: E402

try:
    import pytesseract

    from .tesseract_dependency import locate_tesseract

    # 不能只靠 PATH 找 tesseract.exe——实测在 Windows 上，装完 Tesseract 之后 PATH
    # 没刷新成功是常见情况，这里直接把找到的完整路径显式告诉 pytesseract，绕开 PATH。
    _tesseract_path = locate_tesseract()
    if _tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = _tesseract_path
    _HAS_TESSERACT = _tesseract_path is not None
except ImportError:
    _HAS_TESSERACT = False

_BARCODE_PATTERN = re.compile(
    r"^400(?P<shipment_id>.+?)01(?P<gtin>\d{14})00(?P<box_id>\d{18})30(?P<quantity>\d+)$"
)

# "SINGLE SKU:" 那行印刷文字在 300dpi 渲染下的大致位置（页面是 1200x1800）。
_SKU_OCR_CROP = (550, 1150, 1200, 1550)
_SKU_OCR_DPI = 300


def require_tesseract() -> None:
    """核对流程一开始就调用一下，没装 Tesseract（或者装了但找不到 tesseract.exe）就直接
    报错说清楚原因，不要等跑到一半、所有 SKU 都读不出来才让人自己发现——之前真实踩过这个坑：
    pytesseract 这个 Python 包装好了，但底层的 Tesseract 引擎没装，OCR 每次都静默失败，
    核对结果整批全部"查不到"，界面上看起来像是正常跑完了，其实数据是空的。
    """
    if not _HAS_TESSERACT:
        raise RuntimeError(
            "没有找到 Tesseract（可能是 pytesseract 这个包没装，也可能是包装了但找不到 tesseract.exe），"
            "没法从箱唛上读取 SKU，这次核对流程离不开这一步。"
            "Mac 上装：brew install tesseract；Windows 上装 UB-Mannheim 提供的安装包。"
        )


@dataclass
class BoxLabel:
    page_index: int
    shipment_id: str
    gtin: str
    box_id: str
    quantity: int


def get_page_count(pdf_path: str | Path) -> int:
    doc = fitz.open(str(pdf_path))
    try:
        return doc.page_count
    finally:
        doc.close()


def parse_box_labels(
    pdf_path: str | Path,
    dpi: int = 150,
    max_workers: int | None = None,
    on_page_done: Callable[[], None] | None = None,
    cancel_event: threading.Event | None = None,
) -> tuple[list[BoxLabel], int]:
    """返回 (箱标签列表, 跳过的非箱标签页数)。

    max_workers 不填的话用 CPU 核数——每页单独开一个 fitz 句柄去渲染+解码，
    是因为 fitz.Document 不支持跨进程/跨线程共享同一个句柄，各页之间必须完全独立。

    on_page_done：每解完一页就调用一次（不管这页是箱标签还是被跳过），给界面上报进度用；
    用 as_completed 而不是 map，是因为 map 要按提交顺序把结果攒好才吐出来，没法在每一页
    真正完成的时候就立刻通知外面。

    cancel_event：中途设了这个，会在当前这一批已提交的页面陆续完成的间隙尽快停下来，
    并且把还没跑完的子进程直接杀掉——不能指望 QThread.terminate() 去掐断这个函数，实测
    线程卡在 concurrent.futures 的等待里时，terminate() 经常会直接卡死，比不处理还糟。
    """
    pdf_path = str(pdf_path)
    page_count = get_page_count(pdf_path)
    if page_count == 0:
        return [], 0

    workers = max_workers or os.cpu_count() or 1
    workers = max(1, min(workers, page_count))

    labels: list[BoxLabel] = []
    skipped = 0
    cancelled = False
    pool = ProcessPoolExecutor(max_workers=workers)
    try:
        futures = [pool.submit(_decode_page, pdf_path, i, dpi) for i in range(page_count)]
        for future in as_completed(futures):
            if cancel_event is not None and cancel_event.is_set():
                cancelled = True
                break
            result = future.result()
            if result is None:
                skipped += 1
            else:
                labels.append(result)
            if on_page_done is not None:
                on_page_done()
    finally:
        if cancelled:
            for proc in list(getattr(pool, "_processes", {}).values()):
                if proc.is_alive():
                    proc.kill()
            pool.shutdown(wait=False, cancel_futures=True)
        else:
            pool.shutdown(wait=True)
    return labels, skipped


def _decode_page(pdf_path: str, page_index: int, dpi: int) -> BoxLabel | None:
    """在子进程里跑：单独打开一次 PDF，只渲染、解码这一页。"""
    doc = fitz.open(pdf_path)
    try:
        pix = doc[page_index].get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    finally:
        doc.close()

    for code in _dmtx_decode(img):
        raw = code.data.decode("utf-8", errors="replace")
        m = _BARCODE_PATTERN.match(raw)
        if m:
            return BoxLabel(
                page_index=page_index,
                shipment_id=m.group("shipment_id"),
                gtin=m.group("gtin"),
                box_id=m.group("box_id"),
                quantity=int(m.group("quantity")),
            )
    return None


def ocr_single_sku(pdf_path: str | Path, page_index: int) -> str | None:
    """从箱唛印刷文字里读出 SINGLE SKU 的值。这次核对流程里 GTIN 没法直接对应到任何一份
    表，SKU 文字是唯一能把条码解出来的箱子跟翻译表/发货计划表连起来的字段。同一个 GTIN 只需要
    调用一次（挑一页代表页），不用每页都读。读不出来返回 None，调用方自行决定怎么处理。
    """
    if not _HAS_TESSERACT:
        return None
    try:
        doc = fitz.open(str(pdf_path))
        try:
            pix = doc[page_index].get_pixmap(dpi=_SKU_OCR_DPI)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        finally:
            doc.close()
        crop = img.crop(_SKU_OCR_CROP)
        raw_text = pytesseract.image_to_string(crop)
    except Exception:
        return None

    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    candidates = [line for line in lines if "QUANTITY" not in line.upper() and "SKU" not in line.upper()]
    if not candidates:
        return None

    # 只去掉 OCR 引入的空白噪音，大小写/长破折号这些留给 translation.normalize_sku() 统一处理，
    # 不在这里重复写一遍归一化逻辑
    value = re.sub(r"\s+", "", candidates[-1])
    return value or None
