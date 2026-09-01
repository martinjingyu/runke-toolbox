"""物流仓库模块的界面——目前只有一个功能：发货数量核对（Walmart）。

核对本身要跑几分钟（条码解码是主要耗时），放在后台线程里跑，不然界面会卡死；
线程跑完用信号把结果／错误传回主线程再更新界面，不能在后台线程里直接碰 Qt 控件。
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .walmart_shipment_reconcile import RunReport
from .walmart_shipment_reconcile import run as run_reconcile
from .walmart_shipment_reconcile import write_report_xlsx

_MATCH_COLOR = QColor("#C6E0B4")
_MISMATCH_COLOR = QColor("#F8CBAD")


class _ReconcileWorker(QThread):
    succeeded = Signal(object)  # RunReport
    failed = Signal(str)
    progress = Signal(int, int)  # done, total（页数）

    def __init__(self, pdf_paths: list[str], translation_path: str, plan_path: str, output_dir: str):
        super().__init__()
        self._pdf_paths = pdf_paths
        self._translation_path = translation_path
        self._plan_path = plan_path
        self._output_dir = output_dir

    def run(self):
        try:
            report = run_reconcile(
                self._pdf_paths,
                self._translation_path,
                self._plan_path,
                self._output_dir,
                progress_callback=lambda done, total: self.progress.emit(done, total),
            )
            write_report_xlsx(report, Path(self._output_dir) / "核对结果.xlsx")
        except Exception as exc:  # 任何失败原因都要带回界面告诉用户，不能让线程默默死掉
            self.failed.emit(str(exc))
            return
        self.succeeded.emit(report)


def _file_picker_row(label_text: str, on_browse) -> tuple[QHBoxLayout, QLineEdit]:
    row = QHBoxLayout()
    row.addWidget(QLabel(label_text))
    line_edit = QLineEdit()
    line_edit.setReadOnly(True)
    row.addWidget(line_edit, 1)
    button = QPushButton("浏览…")
    button.clicked.connect(on_browse)
    row.addWidget(button)
    return row, line_edit


class WalmartReconcilePanel(QWidget):
    def __init__(self):
        super().__init__()
        self._pdf_paths: list[str] = []
        self._worker: _ReconcileWorker | None = None

        layout = QVBoxLayout(self)

        title = QLabel("发货数量核对（Walmart）")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        inputs_box = QGroupBox("输入")
        inputs_layout = QVBoxLayout(inputs_box)

        row, self._translation_edit = _file_picker_row("翻译表（WM-SKU → 货号）", self._browse_translation)
        inputs_layout.addLayout(row)

        row, self._plan_edit = _file_picker_row("发货计划表", self._browse_plan)
        inputs_layout.addLayout(row)

        row, self._pdf_edit = _file_picker_row("箱唛 PDF（可多选，一个仓一个文件）", self._browse_pdfs)
        inputs_layout.addLayout(row)

        row, self._output_edit = _file_picker_row("输出目录", self._browse_output)
        inputs_layout.addLayout(row)

        layout.addWidget(inputs_box)

        run_row = QHBoxLayout()
        self._run_button = QPushButton("开始核对")
        self._run_button.clicked.connect(self._start_run)
        run_row.addWidget(self._run_button)
        self._open_output_button = QPushButton("打开输出文件夹")
        self._open_output_button.clicked.connect(self._open_output_folder)
        self._open_output_button.setEnabled(False)
        run_row.addWidget(self._open_output_button)
        run_row.addStretch(1)
        layout.addLayout(run_row)

        self._progress = QProgressBar()
        self._progress.setRange(0, 0)  # 没有精确进度，先用忙碌样式的进度条，跑完再隐藏
        self._progress.hide()
        layout.addWidget(self._progress)

        self._status_label = QLabel("")
        self._status_label.setWordWrap(True)
        layout.addWidget(self._status_label)

        self._table = QTableWidget(0, 8)
        self._table.setHorizontalHeaderLabels(
            ["货号(SKU)", "仓库", "SHIPMENT ID", "计划数量", "实际数量", "箱数", "是否一致", "备注"]
        )
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setStretchLastSection(True)
        self._table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self._table, 1)

    # ---- 文件选择 ----

    def _browse_translation(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择翻译表", "", "Excel 文件 (*.xlsx)")
        if path:
            self._translation_edit.setText(path)

    def _browse_plan(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择发货计划表", "", "Excel 文件 (*.xlsx)")
        if path:
            self._plan_edit.setText(path)

    def _browse_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择箱唛 PDF（一个仓一个文件）", "", "PDF 文件 (*.pdf)")
        if paths:
            self._pdf_paths = paths
            names = "、".join(Path(p).name for p in paths)
            self._pdf_edit.setText(f"已选 {len(paths)} 个文件：{names}")

    def _browse_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self._output_edit.setText(path)

    # ---- 运行 ----

    def _start_run(self):
        translation_path = self._translation_edit.text().strip()
        plan_path = self._plan_edit.text().strip()
        output_dir = self._output_edit.text().strip()

        missing = []
        if not translation_path:
            missing.append("翻译表")
        if not plan_path:
            missing.append("发货计划表")
        if not self._pdf_paths:
            missing.append("箱唛 PDF")
        if not output_dir:
            missing.append("输出目录")
        if missing:
            QMessageBox.warning(self, "缺少输入", "还没选：" + "、".join(missing))
            return

        self._run_button.setEnabled(False)
        self._open_output_button.setEnabled(False)
        self._progress.setRange(0, 0)  # 页数还没数出来之前，先显示忙碌样式，第一次进度回调后会切成百分比
        self._progress.show()
        self._status_label.setText("正在处理……箱唛页数多的话可能要几分钟，请耐心等待，不要关闭窗口。")
        self._table.setRowCount(0)

        self._worker = _ReconcileWorker(self._pdf_paths, translation_path, plan_path, output_dir)
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
        self._status_label.setText(f"正在解析箱唛条码：{done}/{total} 页……")

    def _on_success(self, report: RunReport):
        self._progress.hide()
        self._run_button.setEnabled(True)
        self._open_output_button.setEnabled(True)

        mismatched = [r for r in report.results if not r.match]
        status = (
            f"完成。共 {len(report.results)} 条，{len(mismatched)} 条不一致，"
            f"拆出 {len(report.split_pdf_paths)} 份 PDF。"
        )
        if report.unresolved_gtins:
            status += f" 有 {len(report.unresolved_gtins)} 个 GTIN 没能读出/翻译出 SKU，用原始文字或 GTIN 命名了拆分文件。"
        self._status_label.setText(status)

        self._table.setRowCount(len(report.results))
        for row, item in enumerate(report.results):
            note = "" if item.in_plan else "发货计划表里查不到这个 SHIPMENT ID + 货号的组合"
            values = [
                item.sku,
                item.warehouse,
                item.shipment_id,
                str(item.planned_quantity),
                str(item.actual_quantity),
                str(item.box_count),
                "一致" if item.match else "不一致",
                note,
            ]
            bg = _MATCH_COLOR if item.match else _MISMATCH_COLOR
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                cell.setBackground(bg)
                self._table.setItem(row, col, cell)

    def _on_failure(self, message: str):
        self._progress.hide()
        self._run_button.setEnabled(True)
        self._status_label.setText("核对失败，见弹窗说明。")
        QMessageBox.critical(self, "核对失败", message)

    def _open_output_folder(self):
        output_dir = self._output_edit.text().strip()
        if output_dir:
            QDesktopServices.openUrl(QUrl.fromLocalFile(output_dir))


def build_panel() -> QWidget:
    """物流仓库这个部门的入口——先看到工具列表，点进去才是具体某个工具的界面。
    以后物流仓库加新工具，在这个列表里加一项 ToolInfo 就行，不用改 HubWidget。"""
    from core.hub_widget import HubWidget, ToolInfo

    tools = [
        ToolInfo(
            id="walmart_shipment_reconcile",
            name="发货数量核对（Walmart）",
            description="核对箱唛实际发货数量和发货计划表是否一致，并按 SKU 拆分箱唛 PDF",
            build_panel=WalmartReconcilePanel,
        ),
    ]
    return HubWidget(tools)
