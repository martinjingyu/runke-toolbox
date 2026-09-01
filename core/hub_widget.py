"""部门模块内部的"工具列表 -> 点进去具体功能"导航组件。

一个部门通常会有好几个小工具（物流仓库现在只有一个，以后会加），选中这个部门时应该先看到
有哪些工具可以用，点了哪个才进那个工具的页面，点"返回"能回到列表——而不是选中部门就直接
钻进某一个工具的界面。这个交互模式抽成一个通用组件，哪个部门模块需要同样的结构直接复用，
不用每个模块自己重写一遍。

每个工具可以声明自己需要的额外依赖（见 core/dependency.py）——第一次点开时才检查/安装，
不是软件一启动就把所有工具的依赖全装一遍。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from PySide6.QtCore import QEventLoop, Qt, QThread, Signal
from PySide6.QtWidgets import (
    QLabel,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.dependency import Dependency


@dataclass
class ToolInfo:
    id: str
    name: str
    description: str
    build_panel: Callable[[], QWidget]
    dependencies: list[Dependency] = field(default_factory=list)


class _DependencyInstallWorker(QThread):
    progress = Signal(str)
    finished_ok = Signal()
    failed = Signal(str)

    def __init__(self, deps: list[Dependency]):
        super().__init__()
        self._deps = deps

    def run(self):
        try:
            for dep in self._deps:
                self.progress.emit(f"正在安装 {dep.name} ...")
                dep.install(lambda msg: self.progress.emit(msg))
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.finished_ok.emit()


class HubWidget(QWidget):
    def __init__(self, tools: list[ToolInfo]):
        super().__init__()
        self._tools = tools
        self._panels: dict[str, QWidget] = {}
        self._tool_widgets: dict[str, QWidget] = {}  # 工具自己的界面（不含返回按钮那层容器）

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
            if not self._ensure_dependencies_ready(tool):
                return  # 用户不装，或者装失败了——留在列表页，不打开这个工具

            container = QWidget()
            container_layout = QVBoxLayout(container)
            container_layout.setContentsMargins(12, 12, 12, 12)

            back_button = QPushButton("← 返回")
            back_button.setFixedWidth(90)
            back_button.clicked.connect(lambda: self._stack.setCurrentWidget(self._list_page))
            container_layout.addWidget(back_button)

            tool_widget = tool.build_panel()
            container_layout.addWidget(tool_widget)

            self._tool_widgets[tool.id] = tool_widget
            self._panels[tool.id] = container
            self._stack.addWidget(container)

        self._stack.setCurrentWidget(self._panels[tool.id])

    def _ensure_dependencies_ready(self, tool: ToolInfo) -> bool:
        missing = [dep for dep in tool.dependencies if not dep.is_installed()]
        if not missing:
            return True

        names = "、".join(d.name for d in missing)
        reply = QMessageBox.question(
            self,
            "需要安装额外组件",
            f"「{tool.name}」这个功能需要先安装：{names}\n\n现在安装吗？（只需要装一次，装好之后不用再装）",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return False

        progress = QProgressDialog("准备安装……", "", 0, 0, self)
        progress.setWindowTitle("正在安装")
        progress.setCancelButton(None)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.show()

        worker = _DependencyInstallWorker(missing)
        loop = QEventLoop()
        result = {"ok": False, "error": ""}

        worker.progress.connect(progress.setLabelText)

        def on_ok():
            result["ok"] = True
            loop.quit()

        def on_failed(msg: str):
            result["error"] = msg
            loop.quit()

        worker.finished_ok.connect(on_ok)
        worker.failed.connect(on_failed)
        worker.start()
        loop.exec()
        worker.wait()
        progress.close()

        if not result["ok"]:
            QMessageBox.critical(
                self,
                "安装失败",
                f"没能装好需要的组件：{result['error']}\n\n可以稍后重新点这个功能再试一次。",
            )
            return False
        return True

    def stop_running_tasks(self) -> None:
        # 点"返回"只是切换显示的页面，之前打开过的工具如果有后台任务还在跑（用户没等它跑完
        # 就点返回了），并不会被停掉——所以这里要把每一个打开过的工具都问一遍，不能只看
        # 当前显示的那个。
        for widget in self._tool_widgets.values():
            stop = getattr(widget, "stop_running_tasks", None)
            if callable(stop):
                stop()
