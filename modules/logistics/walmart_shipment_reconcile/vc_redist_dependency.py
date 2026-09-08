"""Microsoft Visual C++ Redistributable (x64) 的检测/安装。

不是 pip 包，是系统运行库，装法跟 core/dependency.py 里 pip_package() 那套不一样，所以
单独写（跟 tesseract_dependency.py 是一个套路）。

为什么需要它：pylibdmtx 装的 libdmtx-64.dll 是用 VC++ 编译的，这台机器没装过 VC++
Redistributable 的话，import pylibdmtx 会在 ctypes.LoadLibrary 那一步报
"FileNotFoundError: Could not find module ... (or one of its dependencies)"——
文件本身在，缺的是它依赖的 vcruntime140.dll/msvcp140.dll，报错信息很容易让人误以为是
pylibdmtx 没装好，其实是这个运行库缺失，跟这个模块要不要梯子没关系。
"""
from __future__ import annotations

import ctypes
import os
import subprocess
import sys
import tempfile
import urllib.request
from typing import Callable

from core.dependency import Dependency

# 微软官方短链，长期稳定指向最新的 VC++ 2015-2022 x64 运行库，国内一般不用梯子也能下。
_INSTALLER_URL = "https://aka.ms/vs/17/release/vc_redist.x64.exe"


def _is_installed() -> bool:
    if sys.platform != "win32":
        return True  # 这个依赖只在 Windows 上有意义
    for dll_name in ("vcruntime140.dll", "msvcp140.dll"):
        try:
            ctypes.WinDLL(dll_name)
        except OSError:
            return False
    return True


def _install(report: Callable[[str], None]) -> None:
    if sys.platform != "win32":
        raise RuntimeError("这台不是 Windows，不需要装 VC++ Redistributable")

    report("正在下载 VC++ Redistributable 安装包……")
    installer_path = os.path.join(tempfile.gettempdir(), "vc_redist.x64.exe")
    urllib.request.urlretrieve(_INSTALLER_URL, installer_path)

    report("正在安装（静默安装，可能需要几十秒）……")
    subprocess.run([installer_path, "/install", "/quiet", "/norestart"], check=True)

    try:
        os.remove(installer_path)
    except OSError:
        pass

    if not _is_installed():
        raise RuntimeError("安装完成了，但还是检测不到运行库——可能需要重启电脑后再试一次")


def vc_redist() -> Dependency:
    return Dependency(
        name="Microsoft Visual C++ Redistributable（pylibdmtx 解析条码要用到）",
        is_installed=_is_installed,
        install=_install,
    )
