"""海外仓运营部门的入口——列出这个部门有哪些工具，点了哪个才进那个工具的界面。

跟 modules/logistics/hub.py 是同一套模式：这个文件本身只 import PySide6 和 core 里的通用
组件，不 import 具体工具用到的重量级库——那些是每个工具自己的事，只有真的点开它才会被
检查/安装/import（见 core/hub_widget.py）。
"""
from __future__ import annotations

from PySide6.QtWidgets import QWidget

from core.dependency import pip_package
from core.hub_widget import HubWidget, ToolInfo


def _build_sales_import_panel() -> QWidget:
    from .sales_import.panel import SalesImportPanel

    return SalesImportPanel()


def build_panel() -> QWidget:
    tools = [
        ToolInfo(
            id="sales_import",
            name="海外仓销量汇总表导入",
            description="把 ERP 导出 + CastleGate 平台 CSV 导出的当天销量，写进《US库存销售明细表》对应仓库的每日销量列",
            build_panel=_build_sales_import_panel,
            dependencies=[
                pip_package("openpyxl", display_name="openpyxl（读写 Excel）"),
            ],
        ),
    ]
    return HubWidget(tools)
