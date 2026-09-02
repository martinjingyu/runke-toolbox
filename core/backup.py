"""写入真实业务文件之前，先在旁边留一份备份——任何"预览一批改动、人工确认了才真的存盘覆盖
原文件"的工具都用得到，不用每个模块自己写一遍。
"""
from __future__ import annotations

import datetime as dt
import shutil
from pathlib import Path


def backup_file(path: str | Path) -> Path:
    """在原文件旁边生成一份带时间戳的备份（同目录、文件名加 .备份-<时间戳> 后缀），
    返回备份文件的路径。备份失败（比如目录没写权限）会直接抛异常，调用方应该在备份失败时
    中止写入，不能带着"这次没备份成"的状态继续往下走。
    """
    path = Path(path)
    timestamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = path.with_name(f"{path.stem}.备份-{timestamp}{path.suffix}")
    shutil.copy2(path, backup_path)
    return backup_path
