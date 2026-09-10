"""采购订单批量导入——这个工具的界面。

跟 shipment_plan_apply 一样要小心：会真的改采购订单汇总表和发货计划汇总表。业务方明确不想
再有"生成预览、人看一遍、再点确认写入"这两步走的操作——所以这里只有一个「开始导入」按钮，
点了就直接扫文件夹、匹配供应商/历史箱规、写进两张表、存盘，一次性跑完，中途不停下来等人看。
安全性靠另外两层保证，不靠人工审这一层：
  1. 写入前该有的校验（订单号去重、表头必须存在、序号/数量等字段的合理性）都在 planner.py
     里，任何一步不对就直接抛异常，整批都不会写进去。
  2. 存盘用 atomic_save_with_backup（见 core/backup.py）：新内容先完整写到旁边的临时文件，
     写成功了才改名换上去、把原文件改名成备份——原文件不会出现"写到一半被写坏"的情况，出问题
     也能马上从备份恢复。

写完之后界面上会显示这一批实际写了哪些行、哪些订单被跳过、哪些行有信息缺失需要人工补——这是
一份"事后报告"，不是"事前审核"，出了问题事后还能看到，只是不会因为要人看一眼而卡住流程。
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

from core.backup import atomic_save_with_backup
from core.diff_preview import format_cell

from .planner import Plan, PlanItem, apply_plan, build_plan
from .supplier_codes import SupplierCodeStore

_SETTINGS_KEY_FOLDER = "purchase_order_import/order_folder"
_SETTINGS_KEY_PURCHASE = "purchase_order_import/purchase_path"
_SETTINGS_KEY_SUMMARY = "purchase_order_import/summary_path"

_RESULT_HEADERS = [
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


class _ImportWorker(QThread):
    # (plan, 采购表备份文件名或None, 发货计划表备份文件名或None)——后两个是 None 就说明
    # 这一批没有能新增的行，压根没碰任何文件（文件夹是空的，或者订单都已经导入过了）。
    succeeded = Signal(object)
    # 进入存盘阶段之前就失败了（文件读取/解析/校验不通过）——两份原文件保证完全没被动过。
    failed = Signal(str)
    # 已经进入存盘阶段才失败——两个备份文件名哪个是 None，就说明哪份原文件完全没被动过。
    save_failed = Signal(str, str, str)
    stage = Signal(str)
    progress = Signal(str, int, int)

    def __init__(self, folder: str, purchase_path: str, summary_path: str, supplier_map: dict[str, str]):
        super().__init__()
        self._folder = folder
        self._purchase_path = purchase_path
        self._summary_path = summary_path
        self._supplier_map = supplier_map

    def run(self):
        try:
            self.stage.emit("正在加载采购订单汇总表文件…")
            purchase_wb = openpyxl.load_workbook(self._purchase_path, data_only=False)
            self.stage.emit("正在加载发货计划汇总表文件…")
            summary_wb = openpyxl.load_workbook(self._summary_path, data_only=False)
            plan = build_plan(
                Path(self._folder),
                purchase_wb.active,
                summary_wb.active,
                self._supplier_map,
                progress_callback=lambda stage, done, total: self.progress.emit(stage, done, total),
            )
            if not plan.items:
                self.succeeded.emit((plan, None, None))
                return

            self.stage.emit("正在写入新增的行…")
            apply_plan(
                plan,
                purchase_wb.active,
                summary_wb.active,
                progress_callback=lambda done, total: self.progress.emit("正在写入新增的行", done, total),
            )
        except Exception as exc:
            self.failed.emit(str(exc))
            return

        # 存盘单独一个 try：apply_plan 之前的任何失败，两份文件都还没被碰过，直接算整批失败
        # 就够了；存盘这一步开始之后，两份文件是分开调用 atomic_save_with_backup 的，可能只有
        # 一份真的换上了新内容，需要分开报告，不能笼统地说"失败了"。
        purchase_backup = None
        summary_backup = None
        try:
            self.stage.emit("正在保存采购订单汇总表…")
            purchase_backup = atomic_save_with_backup(purchase_wb, self._purchase_path)
            self.stage.emit("正在保存发货计划汇总表…")
            summary_backup = atomic_save_with_backup(summary_wb, self._summary_path)
        except Exception as exc:
            self.save_failed.emit(
                str(exc),
                purchase_backup.name if purchase_backup else None,
                summary_backup.name if summary_backup else None,
            )
            return

        self.succeeded.emit((plan, purchase_backup.name, summary_backup.name))


class PurchaseOrderImportPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._settings = QSettings()
        self._supplier_store = SupplierCodeStore(self._settings)
        self._worker: _ImportWorker | None = None

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
            "点「开始导入」会直接扫文件夹、匹配供应商/历史箱规、写进采购订单汇总表和发货计划汇总表、"
            "存盘，一次性跑完，中途不会停下来等确认。存盘时会自动生成这两份文件的备份（新内容先"
            "完整写到旁边的临时文件，写成功了才改名换上去，原文件不会出现「写到一半被写坏」的情况）。"
            "写完之后下面会显示这一批实际新增了哪些行、哪些订单被跳过、哪些行信息缺失需要人工核对——"
            "这是写完之后的报告，不是写之前的审核。"
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
        self._import_button = QPushButton("开始导入")
        self._import_button.clicked.connect(self._start_import)
        run_row.addWidget(self._import_button)
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

        result_box = QGroupBox("这一批新增的行（写入之后的记录，不是写入前的预览）")
        result_layout = QVBoxLayout(result_box)
        self._result_table = QTableWidget(0, len(_RESULT_HEADERS))
        self._result_table.setHorizontalHeaderLabels(_RESULT_HEADERS)
        self._result_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        result_layout.addWidget(self._result_table)
        layout.addWidget(result_box)

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

    # ---- 导入 ----

    def _start_import(self):
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

        # 不展开预览表格让人一行行核对了，但保留这一步最基础的"你选的是不是这几个文件"确认——
        # 防的是选错文件夹/表格这种手滑，不是要求把生成的每一行数据看一遍。
        reply = QMessageBox.question(
            self,
            "开始导入",
            "确定要把这个文件夹里的新订单写进：\n"
            f"  {purchase_path}\n"
            f"  {summary_path}\n\n"
            "点了就直接写入，中途不会再停下来确认。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self._import_button.setEnabled(False)
        self._progress.setRange(0, 0)
        self._progress.show()
        self._skip_text.hide()
        self._status_label.setText("正在处理……订单多的话可能要一会，请耐心等待，不要关闭窗口。")
        self._result_table.setRowCount(0)

        self._worker = _ImportWorker(folder, purchase_path, summary_path, self._supplier_store.mapping())
        self._worker.succeeded.connect(self._on_import_succeeded)
        self._worker.failed.connect(self._on_import_failed)
        self._worker.save_failed.connect(self._on_save_failed)
        self._worker.stage.connect(self._on_stage)
        self._worker.progress.connect(self._on_progress)
        self._worker.start()

    def _on_stage(self, text: str):
        self._progress.setRange(0, 0)
        self._status_label.setText(text)

    def _on_progress(self, stage: str, done: int, total: int):
        if total <= 0:
            return
        if self._progress.maximum() != total:
            self._progress.setRange(0, total)
        self._progress.setValue(done)
        self._status_label.setText(f"{stage}：{done}/{total}")

    def _on_import_succeeded(self, payload):
        plan, purchase_backup_name, summary_backup_name = payload
        self._progress.hide()
        self._import_button.setEnabled(True)

        skip_lines = [f"已跳过：{s.order_no}（{s.source_file}）—— {s.reason}" for s in plan.skipped_orders]
        skip_lines.extend(f"没能处理：{f}" for f in plan.skipped_files)
        if skip_lines:
            self._skip_text.setPlainText("\n".join(skip_lines))
            self._skip_text.show()
        else:
            self._skip_text.hide()

        self._fill_result(plan.items)

        if not plan.items:
            self._status_label.setText("这一批没有能新增的行（文件夹是空的，或者订单都已经导入过了），没有改动任何文件。")
            return

        self._status_label.setText(
            f"已写入 {len(plan.items)} 条记录。备份文件：{purchase_backup_name} / {summary_backup_name}"
        )

    def _fill_result(self, items: list[PlanItem]) -> None:
        table = self._result_table
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

    def _on_import_failed(self, message: str):
        self._progress.hide()
        self._import_button.setEnabled(True)
        self._status_label.setText("导入失败，见弹窗说明；两份原文件都没有被改动。")
        QMessageBox.critical(self, "导入失败", message)

    def _on_save_failed(self, message: str, purchase_backup_name: str | None, summary_backup_name: str | None):
        self._progress.hide()
        self._import_button.setEnabled(True)
        self._status_label.setText("导入失败，见弹窗说明。")
        # 新内容是先完整写到旁边的临时文件、写成功了才改名换上去的（见 core/backup.py 的
        # atomic_save_with_backup），所以这两个备份文件名哪个是 None，就说明哪份原文件完全
        # 没被动过；只有真的换上新内容的那份才会有备份文件名，报错信息按各自的实际情况分开说明。
        lines = [f"存盘失败：{message}", ""]
        if purchase_backup_name:
            lines.append(f"采购订单汇总表已经换上新内容，原文件备份在「{purchase_backup_name}」。")
        else:
            lines.append("采购订单汇总表完全没被改动，不用处理。")
        if summary_backup_name:
            lines.append(f"发货计划汇总表已经换上新内容，原文件备份在「{summary_backup_name}」。")
        else:
            lines.append("发货计划汇总表完全没被改动，不用处理。")
        QMessageBox.critical(self, "导入失败", "\n".join(lines))

    def stop_running_tasks(self):
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(30000)
