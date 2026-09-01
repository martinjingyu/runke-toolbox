"""工具自己声明"我需要什么"，核心壳负责在用户第一次点开这个工具时检查/安装，而不是让
所有人在装软件的时候就把每个工具用到的东西都装一遍。

核心壳自己只依赖 PySide6 + PyYAML（见 requirements.txt），装起来很快；某个部门的某个工具
要读 Excel、要做 OCR 之类，都是那个工具自己的事，只有真的用到的人才会被要求装。
"""
from __future__ import annotations

import importlib.util
import subprocess
import sys
from dataclasses import dataclass
from typing import Callable


@dataclass
class Dependency:
    name: str  # 展示给用户看的名字，比如 "openpyxl（读写 Excel）"
    is_installed: Callable[[], bool]
    install: Callable[[Callable[[str], None]], None]  # 参数是 report(message)，装的过程中用来汇报进度


def pip_package(package: str, import_name: str | None = None, display_name: str | None = None) -> Dependency:
    """最常见的情况——一个能直接 pip install 的包。

    import_name：有些包 pip 名字和 import 名字不一样（比如 pip install pymupdf 但是 import fitz），
    这种情况下要单独传一下，不然 is_installed() 会一直判断成"没装"。
    """
    mod_name = import_name or package

    def is_installed() -> bool:
        return importlib.util.find_spec(mod_name) is not None

    def install(report: Callable[[str], None]) -> None:
        report(f"正在安装 {package} ...")
        kwargs = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        subprocess.run(
            [sys.executable, "-m", "pip", "install", package],
            check=True,
            capture_output=True,
            text=True,
            **kwargs,
        )

    return Dependency(name=display_name or package, is_installed=is_installed, install=install)
