"""采购订单批量导入——这个工具的界面。

跟 shipment_plan_apply 一样要小心：会真的改采购订单汇总表和发货计划汇总表。所以也是
"生成预览"只在内存里跑完（扫文件夹、去重、匹配供应商/历史箱规），人看过没问题再点
"确认写入"，写入前自动备份两份原文件。

这里不用 core/diff_preview.py 那套"变化前/变化后"对比组件——那个组件要求每条记录都有
"改动前"的状态，但这个工具从头到尾都是纯新增（新订单在两张表里之前完全没出现过，没有
"改之前"可比），所以预览就是一张"即将新增哪些行"的表格，外加"哪些订单因为已存在被跳过"
"哪些行有信息缺失需要人工补"的提示列表。
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from PySide6.QtCore import QSettings, Qt, QThread, Signal
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.backup import backup_file
from core.diff_preview import format_cell

from .planner import Plan, PlanItem, apply_plan, build_plan
from .supplier_codes import SupplierCodeStore

_SETTINGS_KEY_FOLDER = "purchase_order_import/order_folder"
_SETTINGS_KEY_PURCHASE = "purchase_order_import/purchase_path"
_SETTINGS_KEY_SUMMARY = "purchase_order_import/summary_path"

_PREVIEW_HEADERS = [
    "订单号", "型号", "产品名称", "数量", "交货日期", "供应商代码",
    "箱容", "箱数", "长", "宽", "高", "毛重", "提示",
]


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


class _SupplierMapDialog(QDialog):
    """供应商映射：订单文件里供应商是全称，两张汇总表里历史上填的都是短代码，这里维护
    "全称 -> 代码"的对照表——改一次立刻存一次，不用另外点保存。
    """

    def __init__(self, store: SupplierCodeStore, parent=None):
        super().__init__(parent)
        self.setWindowTitle("供应商映射")
        self.resize(480, 400)
        self._store = store

        layout = QVBoxLayout(self)
        note = QLabel("订单文件里的供应商全称 -> 两张汇总表里要填的短代码（比如 GH / TZ / SX）。")
        note.setWordWrap(True)
        layout.addWidget(note)

        self._table = QTableWidget(0, 3)
        self._table.setHorizontalHeaderLabels(["供应商全称", "代码", ""])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self._table)

        add_box = QGroupBox("添加映射")
        add_layout = QHBoxLayout(add_box)
        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("供应商全称")
        add_layout.addWidget(self._name_edit)
        self._code_edit = QLineEdit()
        self._code_edit.setPlaceholderText("代码")
        add_layout.addWidget(self._code_edit)
        add_button = QPushButton("添加")
        add_button.clicked.connect(self._add_mapping)
        add_layout.addWidget(add_button)
        layout.addWidget(add_box)

        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_button = QPushButton("关闭")
        close_button.clicked.connect(self.accept)
        close_row.addWidget(close_button)
        layout.addLayout(close_row)

        self._reload_table()

    def _reload_table(self) -> None:
        mapping = self._store.mapping()
        names = sorted(mapping)
        self._table.setRowCount(len(names))
        for i, name in enumerate(names):
            self._table.setItem(i, 0, QTableWidgetItem(name))
            self._table.setItem(i, 1, QTableWidgetItem(mapping[name]))
            del_button = QPushButton("删除")
            del_button.clicked.connect(lambda _checked, n=name: self._delete(n))
            self._table.setCellWidget(i, 2, del_button)
        self._table.resizeColumnsToContents()

    def _add_mapping(self) -> None:
        name = self._name_edit.text().strip()
        code = self._code_edit.text().strip()
        if not name or not code:
            QMessageBox.warning(self, "缺少输入", "供应商全称和代码都要填")
            return
        mapping = self._store.mapping()
        mapping[name] = code
        self._store.set_mapping(mapping)
        self._name_edit.clear()
        self._code_edit.clear()
        self._reload_table()

    def _delete(self, name: str) -> None:
        mapping = self._store.mapping()
        mapping.pop(name, None)
        self._store.set_mapping(mapping)
        self._reload_table()


class _PreviewWorker(QThread):
    succeeded = Signal(object)  # (plan, purchase_wb, summary_wb)
    failed = Signal(str)

    def __init__(self, folder: str, purchase_path: str, summary_path: str, supplier_map: dict[str, str]):
        super().__init__()
        self._folder = folder
        self._purchase_path = purchase_path
        self._summary_path = summary_path
        self._supplier_map = supplier_map

    def run(self):
        try:
            purchase_wb = openpyxl.load_workbook(self._purchase_path, data_only=False)
            summary_wb = openpyxl.load_workbook(self._summary_path, data_only=False)
            plan = build_plan(Path(self._folder), purchase_wb.active, summary_wb.active, self._supplier_map)
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.succeeded.emit((plan, purchase_wb, summary_wb))


class PurchaseOrderImportPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._settings = QSettings()
        self._supplier_store = SupplierCodeStore(self._settings)
        self._worker: _PreviewWorker | None = None

        self._plan: Plan | None = None
        self._purchase_wb = None
        self._summary_wb = None
        self._purchase_path = ""
        self._summary_path = ""

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        outer_layout.addWidget(scroll_area)

        content = QWidget()
        scroll_area.setWidget(content)
        layout = QVBoxLayout(content)

        title = QLabel("采购订单批量导入")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        note = QLabel(
            "会真的修改采购订单汇总表和发货计划汇总表——请先点「生成预览」检查没问题，再点「确认写入」。"
            "确认写入前会自动备份这两份文件的原始版本。新行会照抄表格最后一行的格式和公式（店铺/标签/"
            "DP/CBM/未出货数量等），公式里「引用自己这一行」的部分会自动指向新行，不用业务人员再手动"
            "往下拉一遍。"
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        inputs_box = QGroupBox("输入")
        inputs_layout = QVBoxLayout(inputs_box)
        row, self._folder_edit = _file_picker_row("采购订单文件夹（批量存放订单 .xlsx 的目录）", self._browse_folder)
        inputs_layout.addLayout(row)
        row, self._purchase_edit = _file_picker_row("采购订单汇总表", self._browse_purchase)
        inputs_layout.addLayout(row)
        row, self._summary_edit = _file_picker_row("发货计划汇总表", self._browse_summary)
        inputs_layout.addLayout(row)
        layout.addWidget(inputs_box)

        self._folder_edit.setText(self._settings.value(_SETTINGS_KEY_FOLDER, ""))
        self._purchase_edit.setText(self._settings.value(_SETTINGS_KEY_PURCHASE, ""))
        self._summary_edit.setText(self._settings.value(_SETTINGS_KEY_SUMMARY, ""))

        settings_row = QHBoxLayout()
        supplier_map_button = QPushButton("供应商映射…")
        supplier_map_button.clicked.connect(self._open_supplier_map)
        settings_row.addWidget(supplier_map_button)
        settings_row.addStretch(1)
        layout.addLayout(settings_row)

        run_row = QHBoxLayout()
        self._preview_button = QPushButton("生成预览")
        self._preview_button.clicked.connect(self._start_preview)
        run_row.addWidget(self._preview_button)
        self._confirm_button = QPushButton("确认写入")
        self._confirm_button.setEnabled(False)
        self._confirm_button.clicked.connect(self._confirm_write)
        run_row.addWidget(self._confirm_button)
        run_row.addStretch(1)
        layout.addLayout(run_row)

        self._progress = QProgressBar()
        self._progress.setRange(0, 0)
        self._progress.hide()
        layout.addWidget(self._progress)

        self._status_label = QLabel("")
        self._status_label.setWordWrap(True)
        layout.addWidget(self._status_label)

        self._skip_text = QTextEdit()
        self._skip_text.setReadOnly(True)
        self._skip_text.hide()
        self._skip_text.setMaximumHeight(120)
        layout.addWidget(self._skip_text)

        preview_box = QGroupBox("即将新增的行")
        preview_layout = QVBoxLayout(preview_box)
        self._preview_table = QTableWidget(0, len(_PREVIEW_HEADERS))
        self._preview_table.setHorizontalHeaderLabels(_PREVIEW_HEADERS)
        self._preview_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        preview_layout.addWidget(self._preview_table)
        layout.addWidget(preview_box)

    # ---- 文件选择 ----

    def _browse_folder(self):
        start_dir = self._folder_edit.text() or ""
        path = QFileDialog.getExistingDirectory(self, "选择采购订单文件夹", start_dir)
        if path:
            self._folder_edit.setText(path)
            self._settings.setValue(_SETTINGS_KEY_FOLDER, path)

    def _browse_purchase(self):
        start_dir = str(Path(self._purchase_edit.text()).parent) if self._purchase_edit.text() else ""
        path, _ = QFileDialog.getOpenFileName(self, "选择采购订单汇总表", start_dir, "Excel 文件 (*.xlsx)")
        if path:
            self._purchase_edit.setText(path)
            self._settings.setValue(_SETTINGS_KEY_PURCHASE, path)

    def _browse_summary(self):
        start_dir = str(Path(self._summary_edit.text()).parent) if self._summary_edit.text() else ""
        path, _ = QFileDialog.getOpenFileName(self, "选择发货计划汇总表", start_dir, "Excel 文件 (*.xlsx)")
        if path:
            self._summary_edit.setText(path)
            self._settings.setValue(_SETTINGS_KEY_SUMMARY, path)

    def _open_supplier_map(self):
        dialog = _SupplierMapDialog(self._supplier_store, self)
        dialog.exec()

    # ---- 预览 ----

    def _start_preview(self):
        folder = self._folder_edit.text().strip()
        purchase_path = self._purchase_edit.text().strip()
        summary_path = self._summary_edit.text().strip()

        missing = []
        if not folder:
            missing.append("采购订单文件夹")
        if not purchase_path:
            missing.append("采购订单汇总表")
        if not summary_path:
            missing.append("发货计划汇总表")
        if missing:
            QMessageBox.warning(self, "缺少输入", "还没选：" + "、".join(missing))
            return

        self._preview_button.setEnabled(False)
        self._confirm_button.setEnabled(False)
        self._progress.setRange(0, 0)
        self._progress.show()
        self._skip_text.hide()
        self._status_label.setText("正在处理……订单多的话可能要一会，请耐心等待，不要关闭窗口。")
        self._preview_table.setRowCount(0)

        self._purchase_path = purchase_path
        self._summary_path = summary_path

        self._worker = _PreviewWorker(folder, purchase_path, summary_path, self._supplier_store.mapping())
        self._worker.succeeded.connect(self._on_preview_succeeded)
        self._worker.failed.connect(self._on_preview_failed)
        self._worker.start()

    def _on_preview_succeeded(self, payload):
        plan, purchase_wb, summary_wb = payload
        self._progress.hide()
        self._preview_button.setEnabled(True)
        self._plan = plan
        self._purchase_wb = purchase_wb
        self._summary_wb = summary_wb

        skip_lines = [f"已跳过：{s.order_no}（{s.source_file}）—— {s.reason}" for s in plan.skipped_orders]
        skip_lines.extend(f"没能处理：{f}" for f in plan.skipped_files)
        if skip_lines:
            self._skip_text.setPlainText("\n".join(skip_lines))
            self._skip_text.show()
        else:
            self._skip_text.hide()

        self._fill_preview(plan.items)

        if not plan.items:
            self._status_label.setText("这一批没有能新增的行（文件夹是空的，或者订单都已经导入过了）。")
            self._confirm_button.setEnabled(False)
            return

        self._status_label.setText(f"预览完成，共 {len(plan.items)} 条要新增的记录。确认没问题的话点「确认写入」。")
        self._confirm_button.setEnabled(True)
        self._confirm_button.setText("确认写入")

    def _fill_preview(self, items: list[PlanItem]) -> None:
        table = self._preview_table
        table.setRowCount(len(items))
        for r, item in enumerate(items):
            values = [
                item.order_no,
                item.model,
                item.product_name,
                item.quantity,
                item.delivery_date,
                item.supplier_code or "",
                item.box_capacity,
                item.boxes,
                item.length,
                item.width,
                item.height,
                item.gross_weight,
                "；".join(item.notes),
            ]
            for c, value in enumerate(values):
                cell = QTableWidgetItem(format_cell(value))
                if item.notes and c == len(values) - 1:
                    cell.setBackground(Qt.GlobalColor.yellow)
                table.setItem(r, c, cell)
        table.resizeColumnsToContents()

    def _on_preview_failed(self, message: str):
        self._progress.hide()
        self._preview_button.setEnabled(True)
        self._status_label.setText("预览失败，见弹窗说明。")
        QMessageBox.critical(self, "预览失败", message)

    # ---- 确认写入 ----

    def _confirm_write(self):
        if self._plan is None or not self._plan.items or self._purchase_wb is None:
            return

        reply = QMessageBox.question(
            self,
            "确认写入",
            "确定要把上面预览的新增行写进：\n"
            f"  {self._purchase_path}\n"
            f"  {self._summary_path}\n\n"
            "写入前会先在原文件旁边自动生成一份备份。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            purchase_backup = backup_file(self._purchase_path)
            summary_backup = backup_file(self._summary_path)
        except Exception as exc:
            QMessageBox.critical(self, "备份失败", f"没能先备份原文件，写入已取消：{exc}")
            return

        try:
            apply_plan(self._plan, self._purchase_wb.active, self._summary_wb.active)
            self._purchase_wb.save(self._purchase_path)
            self._summary_wb.save(self._summary_path)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "写入失败",
                f"存盘失败：{exc}\n\n备份文件还在（{purchase_backup} / {summary_backup}），原文件可能已经部分改动，建议手动检查。",
            )
            return

        self._status_label.setText(
            f"已写入。备份文件：{purchase_backup.name} / {summary_backup.name}"
        )
        self._confirm_button.setEnabled(False)
        self._confirm_button.setText("已写入")

    def stop_running_tasks(self):
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(5000)
