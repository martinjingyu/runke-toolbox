"""FBA 标签发货人信息脱敏——这个工具自己的界面。

比"发货数量核对"快很多（没有条码解码、没有 OCR，纯文本处理），但页数多的时候还是放
后台线程跑，不阻塞界面；这里操作都很快，没做协作式取消那套，等它自然跑完就行。
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
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from .redact import RunReport
from .redact import run as run_redact


class _RedactWorker(QThread):
    succeeded = Signal(object)  # RunReport
    failed = Signal(str)
    progress = Signal(int, int)

    def __init__(self, input_dir: str, output_dir: str):
        super().__init__()
        self._input_dir = input_dir
        self._output_dir = output_dir

    def run(self):
        try:
            report = run_redact(
                self._input_dir,
                self._output_dir,
                progress_callback=lambda done, total: self.progress.emit(done, total),
            )
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.succeeded.emit(report)


def _dir_picker_row(label_text: str, on_browse) -> tuple[QHBoxLayout, QLineEdit]:
    row = QHBoxLayout()
    row.addWidget(QLabel(label_text))
    line_edit = QLineEdit()
    line_edit.setReadOnly(True)
    row.addWidget(line_edit, 1)
    button = QPushButton("浏览…")
    button.clicked.connect(on_browse)
    row.addWidget(button)
    return row, line_edit


class FbaLabelRedactPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._worker: _RedactWorker | None = None

        layout = QVBoxLayout(self)

        title = QLabel("FBA 标签发货人信息脱敏")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        inputs_box = QGroupBox("输入")
        inputs_layout = QVBoxLayout(inputs_box)

        row, self._input_edit = _dir_picker_row("箱唛 PDF 所在目录", self._browse_input)
        inputs_layout.addLayout(row)

        row, self._output_edit = _dir_picker_row("输出目录", self._browse_output)
        inputs_layout.addLayout(row)

        layout.addWidget(inputs_box)

        run_row = QHBoxLayout()
        self._run_button = QPushButton("开始处理")
        self._run_button.clicked.connect(self._start_run)
        run_row.addWidget(self._run_button)
        self._open_output_button = QPushButton("打开输出文件夹")
        self._open_output_button.clicked.connect(self._open_output_folder)
        self._open_output_button.setEnabled(False)
        run_row.addWidget(self._open_output_button)
        run_row.addStretch(1)
        layout.addLayout(run_row)

        self._progress = QProgressBar()
        self._progress.hide()
        layout.addWidget(self._progress)

        self._status_label = QLabel("")
        self._status_label.setWordWrap(True)
        layout.addWidget(self._status_label)

        self._log = QPlainTextEdit()
        self._log.setReadOnly(True)
        self._log.setPlaceholderText("需要人工看一下的页面（比如结构跟预期不一样、目的地国家没见过）会列在这里。")
        layout.addWidget(self._log, 1)

    def _browse_input(self):
        path = QFileDialog.getExistingDirectory(self, "选择箱唛 PDF 所在目录")
        if path:
            self._input_edit.setText(path)

    def _browse_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self._output_edit.setText(path)

    def _start_run(self):
        input_dir = self._input_edit.text().strip()
        output_dir = self._output_edit.text().strip()

        missing = []
        if not input_dir:
            missing.append("箱唛 PDF 所在目录")
        if not output_dir:
            missing.append("输出目录")
        if missing:
            QMessageBox.warning(self, "缺少输入", "还没选：" + "、".join(missing))
            return

        self._run_button.setEnabled(False)
        self._open_output_button.setEnabled(False)
        self._progress.setRange(0, 0)
        self._progress.show()
        self._status_label.setText("正在处理……")
        self._log.clear()

        self._worker = _RedactWorker(input_dir, output_dir)
        self._worker.succeeded.connect(self._on_success)
        self._worker.failed.connect(self._on_failure)
        self._worker.progress.connect(self._on_progress)
        self._worker.start()

    def _on_progress(self, done: int, total: int):
        if total <= 0:
            return
        if self._progress.maximum() != total:
            self._progress.setRange(0, total)
        self._progress.setValue(done)
        self._status_label.setText(f"正在处理：{done}/{total} 页……")

    def _on_success(self, report: RunReport):
        self._progress.hide()
        self._run_button.setEnabled(True)
        self._open_output_button.setEnabled(True)

        skipped = report.skipped_results
        self._status_label.setText(
            f"完成。共 {len(report.results)} 页，{report.modified_count} 页已处理，"
            f"{len(skipped)} 页跳过（需要人工看）。输出了 {len(report.output_paths)} 个文件。"
        )
        if skipped:
            lines = [f"{r.file_name} 第 {r.page_index + 1} 页：{r.status}" for r in skipped]
            self._log.setPlainText("\n".join(lines))

    def _on_failure(self, message: str):
        self._progress.hide()
        self._run_button.setEnabled(True)
        self._status_label.setText("处理失败，见弹窗说明。")
        QMessageBox.critical(self, "处理失败", message)

    def _open_output_folder(self):
        output_dir = self._output_edit.text().strip()
        if output_dir:
            QDesktopServices.openUrl(QUrl.fromLocalFile(output_dir))

    def stop_running_tasks(self):
        # 这个工具处理很快（没有条码解码/OCR），关软件时就算有任务在跑，等它自然跑完
        # 也就几秒钟，不需要像"发货数量核对"那样搞协作式取消。
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(5000)
