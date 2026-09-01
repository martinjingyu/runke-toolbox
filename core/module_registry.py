"""发现 modules/ 目录下已注册的部门模块。

每个模块是 modules/<模块id>/ 下的一个包，__init__.py 里放一个 MODULE_INFO 字典
（见 modules/logistics/__init__.py 的例子）。新增部门模块时，照这个格式加一个包即可，
不用改这里的代码。
"""
from __future__ import annotations

import importlib
import pkgutil
from dataclasses import dataclass
from typing import Callable, Optional

import modules as modules_package


@dataclass
class ModuleInfo:
    id: str
    name: str
    description: str
    # 模块自己的界面。不填就用一个"功能开发中"的占位面板（见 core/app.py）。
    build_panel: Optional[Callable[[], object]] = None


def discover_modules() -> list[ModuleInfo]:
    found = []
    for _, module_name, is_pkg in pkgutil.iter_modules(modules_package.__path__):
        if not is_pkg:
            continue
        try:
            mod = importlib.import_module(f"modules.{module_name}")
        except Exception as exc:
            # 某个模块自己的依赖没装好（比如条码解码用到的系统库缺失），不能因为这个把整个
            # 软件启动搞挂——其他部门的模块还要能正常用，所以这里只把这一个模块标成加载失败。
            found.append(ModuleInfo(id=module_name, name=module_name, description=f"加载失败：{exc}"))
            continue
        info = getattr(mod, "MODULE_INFO", None)
        if info is None:
            continue
        found.append(
            ModuleInfo(
                id=info["id"],
                name=info["name"],
                description=info["description"],
                build_panel=info.get("build_panel"),
            )
        )
    return found
