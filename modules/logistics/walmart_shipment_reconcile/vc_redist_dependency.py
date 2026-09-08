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

import os
import subprocess
import sys
import tempfile
import urllib.request
from typing import Callable

from core.dependency import Dependency

# 微软官方短链，长期稳定指向最新的 VC++ 2015-2022 x64 运行库，国内一般不用梯子也能下。
_INSTALLER_URL = "https://aka.ms/vs/17/release/vc_redist.x64.exe"

# 微软官方推荐的检测方式：装了 VC++ 2015-2022 Redistributable 会在这两个注册表位置之一
# （64 位系统上两个都会写）写 Installed=1。之前试过用 ctypes.WinDLL("vcruntime140.dll")
# 判断，结果是假阳性——Python 解释器自己进程里就已经加载了同名 DLL，哪怕系统级根本没装
# 这个运行库，探测也会误判成"已安装"，导致该弹的安装提示没弹出来，直接崩在真正用到
# libdmtx-64.dll 依赖的地方。查注册表才是可靠的判断方式。
_REGISTRY_KEY_PATHS = [
    r"SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64",
    r"SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X64",
]


def _is_installed() -> bool:
    if sys.platform != "win32":
        return True  # 这个依赖只在 Windows 上有意义

    import winreg

    for key_path in _REGISTRY_KEY_PATHS:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                installed, _ = winreg.QueryValueEx(key, "Installed")
                if installed == 1:
                    return True
        except OSError:
            continue
    return False


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
