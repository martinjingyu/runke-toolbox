"""核心壳：负责启动窗口、加载配置、列出已注册的部门模块。

具体部门功能的界面由各模块自己提供；框架阶段先给每个模块一个占位面板。
"""
from __future__ import annotations

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


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
