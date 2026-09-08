"""入库标签 PDF 拆分——这个工具自己的界面。

跟 fba_label_redact 一样是纯文本/PDF 拼页操作，没有条码解码、没有 OCR，实测很快，放后台
线程跑但不需要进度条/协作式取消。
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

from .splitter import OUTPUT_DIR_NAME, SplitReport
from .splitter import run as run_split


class _SplitWorker(QThread):
    succeeded = Signal(object)  # SplitReport
    failed = Signal(str)

    def __init__(self, label_pdf_path: str, shipping_plan_path: str):
        super().__init__()
        self._label_pdf_path = label_pdf_path
        self._shipping_plan_path = shipping_plan_path

    def run(self):
        try:
            report = run_split(self._label_pdf_path, self._shipping_plan_path)
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.succeeded.emit(report)


def _file_picker_row(label_text: str, filter_text: str, on_pick) -> tuple[QHBoxLayout, QLineEdit]:
    row = QHBoxLayout()
    row.addWidget(QLabel(label_text))
    line_edit = QLineEdit()
    line_edit.setReadOnly(True)
    row.addWidget(line_edit, 1)
    button = QPushButton("浏览…")
    button.clicked.connect(lambda: on_pick(line_edit))
    row.addWidget(button)
    return row, line_edit


class WarehouseLabelSplitPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._worker: _SplitWorker | None = None
        self._output_dir: Path | None = None

        layout = QVBoxLayout(self)

        title = QLabel("CA1 入库标签 PDF 拆分")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        inputs_box = QGroupBox("输入")
        inputs_layout = QVBoxLayout(inputs_box)

        row, self._label_pdf_edit = _file_picker_row("入库标签 PDF（比如 CA1.pdf）", "PDF 文件 (*.pdf)", self._browse_label_pdf)
        inputs_layout.addLayout(row)

        row, self._plan_edit = _file_picker_row("发货计划表", "Excel 文件 (*.xlsx *.xlsm)", self._browse_plan)
        inputs_layout.addLayout(row)

        hint = QLabel(
            "按发货计划表里「仓库含 CA1、状态=未发货」的标签，从入库标签 PDF 里抽出对应箱子的页面，"
            "按 厂商代号_标签_箱数合计.pdf 命名，存到标签 PDF 所在目录下新建的「output」文件夹。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray;")
        inputs_layout.addWidget(hint)

        layout.addWidget(inputs_box)

        run_row = QHBoxLayout()
        self._run_button = QPushButton("开始拆分")
        self._run_button.clicked.connect(self._start_run)
        run_row.addWidget(self._run_button)
        self._open_output_button = QPushButton("打开输出文件夹")
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
        self._log.setPlaceholderText("拆分结果和需要人工看一下的标签（比如两边对不上、工厂不一致）会列在这里。")
        layout.addWidget(self._log, 1)

    def _browse_label_pdf(self, line_edit: QLineEdit):
        path, _ = QFileDialog.getOpenFileName(self, "选择入库标签 PDF", "", "PDF 文件 (*.pdf)")
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
            missing.append("入库标签 PDF")
        if not plan_path:
            missing.append("发货计划表")
        if missing:
            QMessageBox.warning(self, "缺少输入", "还没选：" + "、".join(missing))
            return

        self._run_button.setEnabled(False)
        self._open_output_button.setEnabled(False)
        self._status_label.setText("正在处理……")
        self._log.clear()

        self._output_dir = Path(label_pdf_path).parent / OUTPUT_DIR_NAME
        self._worker = _SplitWorker(label_pdf_path, plan_path)
        self._worker.succeeded.connect(self._on_success)
        self._worker.failed.connect(self._on_failure)
        self._worker.start()

    def _on_success(self, report: SplitReport):
        self._run_button.setEnabled(True)
        self._open_output_button.setEnabled(bool(report.outputs))

        self._status_label.setText(f"完成。拆出了 {len(report.outputs)} 个文件，{len(report.notes)} 条需要人工看一下的提示。")

        lines = [
            f"{o.output_path.name}：{o.box_count} 个箱子"
            for o in sorted(report.outputs, key=lambda o: o.label)
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
