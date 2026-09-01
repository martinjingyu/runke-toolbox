"""部门模块内部的"工具列表 -> 点进去具体功能"导航组件。

一个部门通常会有好几个小工具（物流仓库现在只有一个，以后会加），选中这个部门时应该先看到
有哪些工具可以用，点了哪个才进那个工具的页面，点"返回"能回到列表——而不是选中部门就直接
钻进某一个工具的界面。这个交互模式抽成一个通用组件，哪个部门模块需要同样的结构直接复用，
不用每个模块自己重写一遍。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel, QPushButton, QStackedWidget, QVBoxLayout, QWidget


@dataclass
class ToolInfo:
    id: str
    name: str
    description: str
    build_panel: Callable[[], QWidget]


class HubWidget(QWidget):
    def __init__(self, tools: list[ToolInfo]):
        super().__init__()
        self._tools = tools
        self._panels: dict[str, QWidget] = {}

        self._stack = QStackedWidget()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(self._stack)

        self._list_page = self._build_list_page()
        self._stack.addWidget(self._list_page)
        self._stack.setCurrentWidget(self._list_page)

    def _build_list_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(10)

        if not self._tools:
            layout.addWidget(QLabel("这个部门还没有可用的工具。"))
            return page

        for tool in self._tools:
            card = QPushButton(f"{tool.name}\n{tool.description}")
            card.setMinimumHeight(56)
            card.setCursor(Qt.CursorShape.PointingHandCursor)
            card.setStyleSheet(
                "QPushButton { text-align: left; padding: 12px 16px; font-size: 13px; }"
            )
            card.clicked.connect(lambda checked=False, t=tool: self._open_tool(t))
            layout.addWidget(card)

        return page

    def _open_tool(self, tool: ToolInfo) -> None:
        if tool.id not in self._panels:
            container = QWidget()
            container_layout = QVBoxLayout(container)
            container_layout.setContentsMargins(12, 12, 12, 12)

            back_button = QPushButton("← 返回")
            back_button.setFixedWidth(90)
            back_button.clicked.connect(lambda: self._stack.setCurrentWidget(self._list_page))
            container_layout.addWidget(back_button)
            container_layout.addWidget(tool.build_panel())

            self._panels[tool.id] = container
            self._stack.addWidget(container)

        self._stack.setCurrentWidget(self._panels[tool.id])
