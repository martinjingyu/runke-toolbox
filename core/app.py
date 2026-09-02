"""核心壳：负责启动窗口、加载配置、列出已注册的部门模块。

具体部门功能的界面由各模块自己提供；框架阶段先给每个模块一个占位面板。
"""
from __future__ import annotations

import os
import sys

from PySide6.QtWidgets import (
    QApplication,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QSplitter,
    QStackedWidget,
    QWidget,
)

from core.config import load_config
from core.module_registry import discover_modules
from core.storage import get_storage


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("闰科内部工具")
        self.resize(900, 600)

        self.config = load_config()
        self.storage = get_storage(self.config)
        self.modules = discover_modules()

        splitter = QSplitter()
        self.setCentralWidget(splitter)

        self.module_list = QListWidget()
        self.pages = QStackedWidget()

        for module in self.modules:
            self.module_list.addItem(QListWidgetItem(module.name))
            if module.build_panel is not None:
                panel = module.build_panel()
            else:
                panel = QLabel(f"「{module.name}」\n\n{module.description}\n\n（功能开发中）")
                panel.setStyleSheet("padding: 24px; font-size: 14px;")
            self.pages.addWidget(panel)

        self.module_list.currentRowChanged.connect(self.pages.setCurrentIndex)
        if self.modules:
            self.module_list.setCurrentRow(0)

        splitter.addWidget(self.module_list)
        splitter.addWidget(self.pages)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([180, 720])

    def closeEvent(self, event):
        # 关窗口时，如果哪个工具还有后台任务在跑（比如核对还没跑完），把它强制停掉——
        # 不然后台线程/子进程还活着会导致进程赖着不退出。每个模块面板只要实现了
        # stop_running_tasks() 就会被调用，没实现的（比如占位面板）直接跳过。
        for i in range(self.pages.count()):
            widget = self.pages.widget(i)
            stop = getattr(widget, "stop_running_tasks", None)
            if callable(stop):
                stop()
        event.accept()


def main():
    app = QApplication(sys.argv)
    # QSettings 靠这两个名字决定存到哪（Windows 是注册表，Mac/Linux 是一个 ini 文件）——
    # 不设的话每次运行用的位置可能不一致，记住的东西（比如上次选过的文件路径）就会找不到。
    app.setOrganizationName("runke")
    app.setApplicationName("runke-toolbox")
    window = MainWindow()
    window.show()
    exit_code = app.exec()
    # 正常的 sys.exit() 会等 Python 解释器自己收尾，如果有什么线程/子进程没清理干净，
    # 进程可能会卡住不退出。这里直接硬退出，保证关窗口=进程真的结束。
    os._exit(exit_code)
