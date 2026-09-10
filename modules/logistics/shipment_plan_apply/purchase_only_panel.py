"""采购订单分摊更新——这个工具的界面。

跟"发货计划自动更新"是同一套输入（同样读运营提供的发货计划表、同样先查在售产品信息总表
翻译成货号、同样按采购日期从早到晚分摊），但只写采购订单汇总表，完全不碰发货计划汇总表——
发货计划汇总表那边要按 ZD 拆成一条条待发货记录，这个工具不需要那一层，同一个货号不管这一批
里分了几个不同的 ZD，最终都是汇总累加进采购订单汇总表同一个日期列（见 planner.py 的
apply_plan_purchase_only）。

不做预览、不弹二次确认——点一下「写入」就直接跑完整个流程，写入前仍然会先自动备份原文件，
这是唯一保留的安全网。如果这一批发货计划里有数据本身的问题（SKU 查不到货号、解析报错……），
在写入之前就会整批拦下来；"余量不够"是唯一的例外——只把那一行跳过（不写、不算阻塞），其它
行照常写入，跳过的行汇总成一张异常数据表格存在第一份发货计划表旁边（见 planner.py 的
write_skipped_items_report）。分摊时按亚马逊>沃尔玛>海外仓的顺序处理，真缺货的话缺口优先
落在海外仓身上（见 planner.py 的 _TEMPLATE_PRIORITY）。
界面结构跟"发货计划自动更新"保持一致，只是少了发货计划汇总表那一栏输入。
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import openpyxl
from PySide6.QtCore import QDate, QSettings, Qt, QThread, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
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

from .planner import apply_plan_purchase_only, build_plan, write_skipped_items_report
from .product_lookup import load_product_lookup
from .purchase_book import PurchaseBook
from .shipment_templates import PlanLine, list_sheet_names, parse_shipment_plan

_TEMPLATE_LABELS = {"walmart": "沃尔玛", "amazon": "亚马逊", "overseas": "海外仓"}
_LABEL_TO_TEMPLATE = {v: k for k, v in _TEMPLATE_LABELS.items()}
_AUTO_LABEL = "自动识别"

# 这两张表长期维护、路径基本不会变，记住上次选过的路径——跟"发货计划自动更新"分开存一份，
# 不共用同一个 key，避免这两个工具互相覆盖对方记的路径（这两个工具选的往往是不同的文件）。
_SETTINGS_KEY_PRODUCT = "purchase_allocation_apply/product_path"
_SETTINGS_KEY_PURCHASE = "purchase_allocation_apply/purchase_path"


def _next_wednesday(today: dt.date | None = None) -> dt.date:
    today = today or dt.date.today()
    days_ahead = (2 - today.weekday()) % 7  # Monday=0 ... Wednesday=2
    return today + dt.timedelta(days=days_ahead)


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


class _PlanFileEntry:
    def __init__(self, path: str, sheet_names: list[str], detected: str | None):
        self.path = path
        self.sheet_names = sheet_names
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItems(sheet_names)
        self.template_combo = QComboBox()
        self.template_combo.addItem(_AUTO_LABEL)
        self.template_combo.addItems(list(_TEMPLATE_LABELS.values()))
        if detected in _TEMPLATE_LABELS:
            self.template_combo.setCurrentText(_TEMPLATE_LABELS[detected])

    @property
    def sheet_name(self) -> str:
        return self.sheet_combo.currentText()

    @property
    def template_type(self) -> str | None:
        label = self.template_combo.currentText()
        return _LABEL_TO_TEMPLATE.get(label)


class _WriteWorker(QThread):
    blocked = Signal(list)
    backup_failed = Signal(str)
    save_failed = Signal(str, str)  # 报错信息, 采购订单汇总表备份文件名
    succeeded = Signal(str, int, str)  # 采购订单汇总表备份文件名, 跳过条数, 异常表格路径（没有就是空字符串）
    failed = Signal(str)
    stage = Signal(str)
    progress = Signal(str, int, int)

    def __init__(
        self,
        product_path: str,
        purchase_path: str,
        plan_files: list[tuple[str, str, str | None]],
        ship_date: dt.date,
    ):
        super().__init__()
        self._product_path = product_path
        self._purchase_path = purchase_path
        self._plan_files = plan_files
        self._ship_date = ship_date

    def run(self):
        try:
            self.stage.emit("正在读取在售产品信息总表…")
            lookup = load_product_lookup(Path(self._product_path))

            all_lines: list[PlanLine] = []
            parse_errors: list[str] = []
            for path, sheet_name, template_type in self._plan_files:
                self.stage.emit(f"正在解析「{Path(path).name}」…")
                parsed = parse_shipment_plan(Path(path), sheet_name, template_type)
                file_label = Path(path).name
                for line in parsed.lines:
                    line.source_file = file_label
                all_lines.extend(parsed.lines)
                parse_errors.extend(f"[{file_label}] {e}" for e in parsed.errors)

            self.stage.emit("正在加载采购订单汇总表…")
            purchase_wb = openpyxl.load_workbook(self._purchase_path, data_only=False)
            purchase_book = PurchaseBook(
                purchase_wb.active,
                progress_callback=lambda done, total: self.progress.emit(
                    "正在加载采购订单汇总表", done, total
                ),
            )

            self.stage.emit("正在校验每一条分摊…")
            plan = build_plan(all_lines, parse_errors, lookup, purchase_book, self._ship_date)

            if plan.has_blocking_errors:
                lines = list(plan.parse_errors)
                for item in plan.items:
                    lines.extend(item.errors)
                self.blocked.emit(lines)
                return

            self.stage.emit("正在写入变化…")
            apply_plan_purchase_only(
                plan,
                purchase_book,
                progress_callback=lambda done, total: self.progress.emit("正在写入变化", done, total),
            )

            self.stage.emit("正在备份原文件…")
            try:
                purchase_backup = backup_file(self._purchase_path)
            except Exception as exc:
                self.backup_failed.emit(str(exc))
                return

            try:
                self.stage.emit("正在存盘采购订单汇总表…")
                purchase_wb.save(self._purchase_path)
            except Exception as exc:
                self.save_failed.emit(str(exc), purchase_backup.name)
                return

            # 异常表格是"锦上添花"的附加产物，不是这次写入成不成功的一部分——文件都已经存盘了，
            # 这一步哪怕出问题也不该把整次操作报成失败，只是异常表格没生成而已。
            report_path = ""
            if plan.skipped_items:
                first_plan_path = Path(self._plan_files[0][0])
                try:
                    written = write_skipped_items_report(plan.skipped_items, first_plan_path)
                    report_path = str(written) if written is not None else ""
                except Exception:
                    report_path = ""

            self.succeeded.emit(purchase_backup.name, len(plan.skipped_items), report_path)
        except Exception as exc:
            self.failed.emit(str(exc))


class PurchaseAllocationApplyPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._plan_entries: list[_PlanFileEntry] = []
        self._worker: _WriteWorker | None = None
        self._settings = QSettings()

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        outer_layout.addWidget(scroll_area)

        content = QWidget()
        scroll_area.setWidget(content)
        layout = QVBoxLayout(content)

        title = QLabel("采购订单分摊更新")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        note = QLabel(
            "只会修改采购订单汇总表，不碰发货计划汇总表——点「写入」直接生效，不会再弹预览确认。"
            "写入前会自动先备份原文件；这一批发货计划有数据问题的话，会在写入前整批拦下来。"
            "分摊按亚马逊>沃尔玛>海外仓的顺序处理，哪一行余量不够会被单独跳过（不算整批出错），"
            "跳过的行会汇总成一张异常数据表格。"
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        inputs_box = QGroupBox("长期维护的两张表")
        inputs_layout = QVBoxLayout(inputs_box)
        row, self._product_edit = _file_picker_row("在售产品信息总表", self._browse_product)
        inputs_layout.addLayout(row)
        row, self._purchase_edit = _file_picker_row("采购订单汇总表", self._browse_purchase)
        inputs_layout.addLayout(row)
        layout.addWidget(inputs_box)

        self._product_edit.setText(self._settings.value(_SETTINGS_KEY_PRODUCT, ""))
        self._purchase_edit.setText(self._settings.value(_SETTINGS_KEY_PURCHASE, ""))

        plan_box = QGroupBox("运营提供的发货计划表（可以一次导入好几份）")
        plan_layout = QVBoxLayout(plan_box)
        add_row = QHBoxLayout()
        add_button = QPushButton("添加文件…")
        add_button.clicked.connect(self._add_plan_files)
        add_row.addWidget(add_button)
        add_row.addStretch(1)
        plan_layout.addLayout(add_row)

        self._plan_table = QTableWidget(0, 4)
        self._plan_table.setHorizontalHeaderLabels(["文件", "选哪个 sheet", "模板类型", ""])
        self._plan_table.horizontalHeader().setStretchLastSection(False)
        self._plan_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        plan_layout.addWidget(self._plan_table)
        layout.addWidget(plan_box)

        date_row = QHBoxLayout()
        date_row.addWidget(QLabel("这次发货日期（采购汇总表写进这一列）"))
        self._date_edit = QDateEdit()
        self._date_edit.setCalendarPopup(True)
        default_date = _next_wednesday()
        self._date_edit.setDate(QDate(default_date.year, default_date.month, default_date.day))
        date_row.addWidget(self._date_edit)
        date_row.addStretch(1)
        layout.addLayout(date_row)

        run_row = QHBoxLayout()
        self._write_button = QPushButton("写入")
        self._write_button.clicked.connect(self._start_write)
        run_row.addWidget(self._write_button)
        run_row.addStretch(1)
        layout.addLayout(run_row)

        self._progress = QProgressBar()
        self._progress.setRange(0, 0)
        self._progress.hide()
        layout.addWidget(self._progress)

        self._status_label = QLabel("")
        self._status_label.setWordWrap(True)
        layout.addWidget(self._status_label)

        self._error_text = QTextEdit()
        self._error_text.setReadOnly(True)
        self._error_text.hide()
        self._error_text.setMaximumHeight(160)
        layout.addWidget(self._error_text)

    # ---- 文件选择 ----

    def _browse_product(self):
        start_dir = str(Path(self._product_edit.text()).parent) if self._product_edit.text() else ""
        path, _ = QFileDialog.getOpenFileName(self, "选择在售产品信息总表", start_dir, "Excel 文件 (*.xlsx)")
        if path:
            self._product_edit.setText(path)
            self._settings.setValue(_SETTINGS_KEY_PRODUCT, path)

    def _browse_purchase(self):
        start_dir = str(Path(self._purchase_edit.text()).parent) if self._purchase_edit.text() else ""
        path, _ = QFileDialog.getOpenFileName(self, "选择采购订单汇总表", start_dir, "Excel 文件 (*.xlsx)")
        if path:
            self._purchase_edit.setText(path)
            self._settings.setValue(_SETTINGS_KEY_PURCHASE, path)

    def _add_plan_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择运营提供的发货计划表", "", "Excel 文件 (*.xlsx)")
        for path in paths:
            try:
                sheet_names = list_sheet_names(Path(path))
                detected = None
                if sheet_names:
                    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                    from .shipment_templates import detect_template_type

                    detected = detect_template_type(wb[sheet_names[0]])
            except Exception as exc:
                QMessageBox.warning(self, "读取失败", f"「{Path(path).name}」读取失败：{exc}")
                continue

            entry = _PlanFileEntry(path, sheet_names, detected)
            self._plan_entries.append(entry)
            self._append_plan_row(entry)

    def _append_plan_row(self, entry: _PlanFileEntry) -> None:
        row = self._plan_table.rowCount()
        self._plan_table.insertRow(row)
        self._plan_table.setItem(row, 0, QTableWidgetItem(Path(entry.path).name))
        self._plan_table.setCellWidget(row, 1, entry.sheet_combo)
        self._plan_table.setCellWidget(row, 2, entry.template_combo)
        remove_button = QPushButton("移除")
        remove_button.clicked.connect(lambda: self._remove_plan_entry(entry))
        self._plan_table.setCellWidget(row, 3, remove_button)
        self._plan_table.resizeColumnsToContents()

    def _remove_plan_entry(self, entry: _PlanFileEntry) -> None:
        if entry not in self._plan_entries:
            return
        idx = self._plan_entries.index(entry)
        self._plan_entries.pop(idx)
        self._plan_table.removeRow(idx)

    # ---- 写入 ----

    def _start_write(self):
        product_path = self._product_edit.text().strip()
        purchase_path = self._purchase_edit.text().strip()

        missing = []
        if not product_path:
            missing.append("在售产品信息总表")
        if not purchase_path:
            missing.append("采购订单汇总表")
        if not self._plan_entries:
            missing.append("发货计划表（至少一份）")
        if missing:
            QMessageBox.warning(self, "缺少输入", "还没选：" + "、".join(missing))
            return

        qdate = self._date_edit.date()
        ship_date = dt.date(qdate.year(), qdate.month(), qdate.day())

        plan_files = [(e.path, e.sheet_name, e.template_type) for e in self._plan_entries]

        self._write_button.setEnabled(False)
        self._progress.setRange(0, 0)
        self._progress.show()
        self._error_text.hide()
        self._status_label.setText("正在处理……表格行数多的话可能要几分钟，请耐心等待，不要关闭窗口。")

        self._worker = _WriteWorker(product_path, purchase_path, plan_files, ship_date)
        self._worker.blocked.connect(self._on_blocked)
        self._worker.backup_failed.connect(self._on_backup_failed)
        self._worker.save_failed.connect(self._on_save_failed)
        self._worker.succeeded.connect(self._on_succeeded)
        self._worker.failed.connect(self._on_failed)
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

    def _on_blocked(self, error_lines: list):
        self._progress.hide()
        self._write_button.setEnabled(True)
        self._status_label.setText(f"这一批有 {len(error_lines)} 处问题，全部列在下面，一个都没写入。")
        self._error_text.setPlainText("\n".join(error_lines))
        self._error_text.show()

    def _on_succeeded(self, purchase_backup_name: str, skipped_count: int, report_path: str):
        self._progress.hide()
        text = f"已写入。备份文件：{purchase_backup_name}"
        if skipped_count > 0:
            text += f"\n有 {skipped_count} 行因为余量不够被跳过，没有写入。"
            text += f"\n异常数据表格：{report_path}" if report_path else "\n（异常数据表格生成失败，跳过的记录只能自己核对）"
        self._status_label.setText(text)
        self._write_button.setText("已写入")

    def _on_failed(self, message: str):
        self._progress.hide()
        self._write_button.setEnabled(True)
        self._status_label.setText("写入失败，见弹窗说明。")
        QMessageBox.critical(self, "写入失败", message)

    def _on_backup_failed(self, message: str):
        self._progress.hide()
        self._write_button.setEnabled(True)
        self._status_label.setText("写入失败，见弹窗说明。")
        QMessageBox.critical(self, "备份失败", f"没能先备份原文件，写入已取消：{message}")

    def _on_save_failed(self, message: str, purchase_backup_name: str):
        self._progress.hide()
        self._write_button.setEnabled(True)
        self._status_label.setText("写入失败，见弹窗说明。")
        QMessageBox.critical(
            self,
            "写入失败",
            f"存盘失败：{message}\n\n"
            f"备份文件还在（{purchase_backup_name}），原文件可能已经部分改动，建议手动检查。",
        )

    def stop_running_tasks(self):
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(30000)
