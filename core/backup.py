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


def atomic_replace_with_backup(path: str | Path, write_new_content) -> Path:
    """跟 atomic_save_with_backup 是同一套"先在别处把新内容完整写出来，成功了才原子改名
    切换"的安全模式，只是不假设新内容是通过 openpyxl 的 wb.save() 产生的——有些场景不能用
    openpyxl 存盘（比如 modules/overseas_warehouse/sales_import 那个场景：目标文件里用了
    openpyxl 完全不认识的 WPS 专有部件，openpyxl 存盘会把这些部件丢光，只能自己手动拼 zip），
    调用方传一个 write_new_content(tmp_path) 的回调，自己决定怎么把新内容写到 tmp_path，
    这里只负责"写失败就不碰原文件、写成功了才原子换上去、原文件挪去当备份"这部分。
    """
    path = Path(path)
    timestamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = path.with_name(f"{path.stem}.备份-{timestamp}{path.suffix}")
    tmp_path = path.with_name(f"{path.stem}.写入中-{timestamp}{path.suffix}")

    write_new_content(tmp_path)  # 失败的话异常直接往外抛，原文件这时候还完全没被碰过

    try:
        if path.exists():
            path.replace(backup_path)
        tmp_path.replace(path)
    except OSError as exc:
        raise OSError(f"改名切换失败，新内容还完整保留在「{tmp_path}」，需要手动处理") from exc

    return backup_path


def atomic_save_with_backup(wb, path: str | Path) -> Path:
    """把 wb 存到 path，但不直接在原文件上覆盖写——先把新内容存到同目录下一个临时文件，
    存成功了才把原文件改名成备份、把临时文件改名成 path 这个名字（改名都是同一块磁盘上的
    目录项操作，不是真的把文件内容再复制一遍，几乎不占时间）。

    这么做省下来的不是时间——wb.save() 该多贵还是多贵，写到临时文件跟直接写原文件是同一份
    开销，两次改名本身可以忽略不计——省下来的是"写到一半就出问题"时的安全性：backup_file()
    +直接 wb.save(path) 这种写法，如果 save() 中途失败（磁盘满/进程被杀/权限问题……），
    原文件本身可能已经被写坏了一部分，备份虽然还在，但原文件已经不是"没被动过"的状态。这里
    改成先在别处把新内容完整写出来，wb.save() 失败的话异常直接抛出去、原文件完全没被碰过，
    调用方不用担心"这次是不是把原文件写坏了"，只有确认新内容整个写完了才做改名切换。

    返回值：备份文件的路径（只有真的换上新内容之后才存在）。
    """
    path = Path(path)
    timestamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = path.with_name(f"{path.stem}.备份-{timestamp}{path.suffix}")
    tmp_path = path.with_name(f"{path.stem}.写入中-{timestamp}{path.suffix}")

    wb.save(tmp_path)  # 失败的话异常直接往外抛，原文件这时候还完全没被碰过

    try:
        if path.exists():
            # Path.rename() 在 Windows 上遇到目标已存在会直接报错（不像 POSIX 的 rename
            # 会原子替换）——统一用 .replace()，两个平台都是"目标存在就替换掉"。
            path.replace(backup_path)
        tmp_path.replace(path)
    except OSError as exc:
        # 改名这一步理论上不会真的失败（backup_path 是新算出来的时间戳文件名，不会已经
        # 存在；tmp_path 刚写完必然存在）——真出现了，新内容还完整地留在 tmp_path 里，
        # 没有丢，把这个路径带在异常信息里，方便人工去把它手动改名成正确的文件名。
        raise OSError(f"改名切换失败，新内容还完整保留在「{tmp_path}」，需要手动处理") from exc

    return backup_path
