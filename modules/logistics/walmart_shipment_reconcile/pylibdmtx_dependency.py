"""pylibdmtx 的检测/安装。

跟 core/dependency.py 里普通的 pip_package() 不一样的地方：pip_package() 的 is_installed()
只用 importlib.util.find_spec() 看包的 .py 文件在不在，这个包能骗过这个检查——pylibdmtx 是
纯 Python 包装，真正干活的是随包附带的 libdmtx-64.dll（Windows 上），只有在真的
`from pylibdmtx.pylibdmtx import decode` 的时候才会去加载这个 DLL。如果这个 DLL 缺失或者
损坏（比如上次安装被杀毒软件拦了、装到一半断网了），find_spec() 照样返回"找到了"，依赖检查
就会误判成"已装好"，用户点开工具那一刻才在 import 阶段崩掉，报 "Could not find module
'libdmtx-64.dll'"——这时候再点一次这个工具，检查还是显示"已装好"，用户没有任何办法自己修，
只能来问人。所以这里的 is_installed() 得真的做一次会触发 DLL 加载的 import，而不是只看文件在不在；
判断成"没装"的话，install() 用 --force-reinstall 强制重新拉一遍包，把可能损坏的 DLL 文件覆盖掉
（普通的 pip install 一看包"已经满足要求"就什么都不做，treat 不了损坏的安装）。
"""
from __future__ import annotations

import subprocess
import sys
from typing import Callable

from core.dependency import Dependency


def _can_load() -> bool:
    try:
        from pylibdmtx.pylibdmtx import decode  # noqa: F401
    except Exception:
        # 可能是包没装（ModuleNotFoundError），也可能是包在但 DLL 加载失败
        # （Windows 上典型是 FileNotFoundError/OSError）——两种都当"没装好"处理
        return False
    return True


def _install(report: Callable[[str], None]) -> None:
    report("正在安装 pylibdmtx ...")
    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    subprocess.run(
        [sys.executable, "-m", "pip", "install", "--force-reinstall", "--no-cache-dir", "pylibdmtx"],
        check=True,
        capture_output=True,
        text=True,
        **kwargs,
    )
    if not _can_load():
        raise RuntimeError(
            "重新安装完成了，但还是加载不了 pylibdmtx 的解码库——"
            "如果这台机器还是报 libdmtx-64 找不到，多半是 Microsoft Visual C++ Redistributable "
            "没装好（这个工具需要的另一项依赖），或者是杀毒软件把 DLL 文件拦下来了，需要手动检查一下"
        )


def pylibdmtx_decoder() -> Dependency:
    return Dependency(
        name="pylibdmtx（解析箱唛条码）",
        is_installed=_can_load,
        install=_install,
    )
