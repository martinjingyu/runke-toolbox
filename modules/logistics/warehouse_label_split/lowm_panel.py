"""LO-WM（Walmart）站点箱唛 PDF 拆分——这个工具自己的界面。

跟 CA1/CG 拆分工具用的是同一套 _file_picker_row 之类的小组件（见 panel.py），但核心逻辑
（lowm_splitter.run）不需要认 SKU，报告的形状也不一样（按厂商汇总，不是按标签），所以单独
写一个面板类，不硬塞进 LabelSplitPanel 里。
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from .lowm_splitter import LowmSplitReport
from .lowm_splitter import run as run_lowm_split
from .panel import _file_picker_row


class _LowmSplitWorker(QThread):
    succeeded = Signal(object)  # LowmSplitReport
    failed = Signal(str)

    def __init__(self, label_pdf_path: str, shipping_plan_path: str):
        super().__init__()
        self._label_pdf_path = label_pdf_path
        self._shipping_plan_path = shipping_plan_path

    def run(self):
        try:
            report = run_lowm_split(self._label_pdf_path, self._shipping_plan_path)
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.succeeded.emit(report)


class LowmSplitPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._worker: _LowmSplitWorker | None = None
        self._output_dir: Path | None = None

        layout = QVBoxLayout(self)

        title_label = QLabel("LO-WM（Walmart）箱唛 PDF 拆分")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title_label)

        inputs_box = QGroupBox("输入")
        inputs_layout = QVBoxLayout(inputs_box)

        row, self._label_pdf_edit = _file_picker_row("站点箱唛 PDF（比如 DFW5s.pdf）", "PDF 文件 (*.pdf)", self._browse_label_pdf)
        inputs_layout.addLayout(row)

        row, self._plan_edit = _file_picker_row("发货计划表", "Excel 文件 (*.xlsx *.xlsm)", self._browse_plan)
        inputs_layout.addLayout(row)

        hint = QLabel(
            "站点代号取自箱唛 PDF 的文件名（第一个「-」之前的部分）。"
            "不需要认 SKU：按发货计划表里「仓库含这个站点代号、状态=未发货」的记录按厂商汇总箱数，"
            "直接按顺序切页给各厂商，在箱唛 PDF 所在目录下按「厂商代号/站点代号」新建两层文件夹（比如「GH/DFW5s」），"
            "文件按「厂商代号 站点代号 箱数箱.pdf」命名。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray;")
        inputs_layout.addWidget(hint)

        layout.addWidget(inputs_box)

        run_row = QHBoxLayout()
        self._run_button = QPushButton("开始拆分")
        self._run_button.clicked.connect(self._start_run)
        run_row.addWidget(self._run_button)
        self._open_output_button = QPushButton("打开所在目录")
        self._open_output_button.clicked.connect(self._open_output_folder)
        self._open_output_button.setEnabled(False)
        run_row.addWidget(self._open_output_button)
        run_row.addStretch(1)
        layout.addLayout(run_row)

        self._status_label = QLabel("")
        self._status_label.setWordWrap(True)
        layout.addWidget(self._status_label)

        self._log = QPlainTextEdit()
        self._log.setReadOnly(True)
        self._log.setPlaceholderText("拆分结果和需要人工看一下的提示（比如页数不够、箱数没对齐）会列在这里。")
        layout.addWidget(self._log, 1)

    def _browse_label_pdf(self, line_edit: QLineEdit):
        path, _ = QFileDialog.getOpenFileName(self, "选择站点箱唛 PDF", "", "PDF 文件 (*.pdf)")
        if path:
            line_edit.setText(path)

    def _browse_plan(self, line_edit: QLineEdit):
        path, _ = QFileDialog.getOpenFileName(self, "选择发货计划表", "", "Excel 文件 (*.xlsx *.xlsm)")
        if path:
            line_edit.setText(path)

    def _start_run(self):
        label_pdf_path = self._label_pdf_edit.text().strip()
        plan_path = self._plan_edit.text().strip()

        missing = []
        if not label_pdf_path:
            missing.append("站点箱唛 PDF")
        if not plan_path:
            missing.append("发货计划表")
        if missing:
            QMessageBox.warning(self, "缺少输入", "还没选：" + "、".join(missing))
            return

        self._run_button.setEnabled(False)
        self._open_output_button.setEnabled(False)
        self._status_label.setText("正在处理……")
        self._log.clear()

        self._output_dir = Path(label_pdf_path).parent
        self._worker = _LowmSplitWorker(label_pdf_path, plan_path)
        self._worker.succeeded.connect(self._on_success)
        self._worker.failed.connect(self._on_failure)
        self._worker.start()

    def _on_success(self, report: LowmSplitReport):
        self._run_button.setEnabled(True)
        self._open_output_button.setEnabled(bool(report.outputs))

        self._status_label.setText(f"完成。拆出了 {len(report.outputs)} 个厂商的文件，{len(report.notes)} 条需要人工看一下的提示。")

        lines = [
            f"{o.output_path.parent.parent.name}/{o.output_path.parent.name}/{o.output_path.name}"
            for o in sorted(report.outputs, key=lambda o: o.factory)
        ]
        if report.notes:
            lines.append("")
            lines.extend(report.notes)
        if lines:
            self._log.setPlainText("\n".join(lines))

    def _on_failure(self, message: str):
        self._run_button.setEnabled(True)
        self._status_label.setText("处理失败，见弹窗说明。")
        QMessageBox.critical(self, "处理失败", message)

    def _open_output_folder(self):
        if self._output_dir is not None:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._output_dir)))

    def stop_running_tasks(self):
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(5000)
