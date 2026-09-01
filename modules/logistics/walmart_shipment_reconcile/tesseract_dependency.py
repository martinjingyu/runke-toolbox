"""Tesseract OCR 的检测/安装。这是一个外部程序，不是 pip 包，装法跟 core/dependency.py
里 pip_package() 那套不一样，所以单独写。

这个文件本身只用标准库（不 import fitz/pylibdmtx/pytesseract 这些），是为了让"物流仓库"
这个部门的工具列表页能便宜地列出这个工具、检查它需要装什么，而不用先把这个工具真正用到的
重量级的库都导入一遍。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from typing import Callable

from core.dependency import Dependency

_WINDOWS_CANDIDATES = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]
_INSTALLER_URL = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe"


def locate_tesseract() -> str | None:
    """找 tesseract.exe 的完整路径，不能只看 PATH 里有没有——之前踩过坑，装完之后当前
    进程/新开的窗口不一定马上认得到 PATH，所以固定装的位置也直接查一遍，更可靠。
    """
    found = shutil.which("tesseract")
    if found:
        return found
    for candidate in _WINDOWS_CANDIDATES:
        if os.path.exists(candidate):
            return candidate
    return None


def _install(report: Callable[[str], None]) -> None:
    if sys.platform != "win32":
        raise RuntimeError("这台不是 Windows，请手动安装：brew install tesseract（Mac）")

    report("正在下载 Tesseract 安装包……")
    installer_path = os.path.join(tempfile.gettempdir(), "tesseract-installer.exe")
    urllib.request.urlretrieve(_INSTALLER_URL, installer_path)

    report("正在安装（Windows 可能会跳出授权提示，请点“是”）……")
    subprocess.run([installer_path, "/S"], check=True)

    try:
        os.remove(installer_path)
    except OSError:
        pass

    if locate_tesseract() is None:
        raise RuntimeError("安装完成了，但还是找不到 tesseract.exe——可能是没有管理员权限，安装被跳过了")


def tesseract_ocr() -> Dependency:
    return Dependency(
        name="Tesseract OCR（读取箱唛上的 SKU 文字）",
        is_installed=lambda: locate_tesseract() is not None,
        install=_install,
    )
