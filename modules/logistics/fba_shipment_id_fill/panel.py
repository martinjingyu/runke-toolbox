"""FBA 货件编号回填——这个工具的界面。

跟 shipment_plan_apply 系列一样：不做预览、不弹二次确认，点一下「写入」直接跑完整个流程，
写入前先自动备份原文件；发现任何问题（数量对不上、目标行已经有 FBA ID、箱容除不尽……）会
在写入前整批拦下来，列在报错框里，一个都不会写进去。写整张几万行的表可能要一点时间，放在
后台线程里跑，不能卡住界面。
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import openpyxl
from PySide6.QtCore import QDate, QSettings, QThread, Signal
from PySide6.QtWidgets import (
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
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.backup import atomic_save_with_backup

from .fba_source import parse_fba_folder
from .plan_matcher import Plan, apply_plan, build_plan
from .report import write_skip_report

_SETTINGS_KEY_PLAN = "fba_shipment_id_fill/plan_path"
_SETTINGS_KEY_FBA_FOLDER = "fba_shipment_id_fill/fba_folder"
_MAIN_SHEET_NAME = "发货计划(2)"


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


class _WriteWorker(QThread):
    blocked = Signal(list)
    failed = Signal(str)
    succeeded = Signal(str, int, str)  # 备份文件名, 跳过条数, 跳过说明表格路径（没有就是空字符串）
    stage = Signal(str)

    def __init__(self, plan_path: str, fba_folder: str, ship_date: dt.date):
        super().__init__()
        self._plan_path = plan_path
        self._fba_folder = fba_folder
        self._ship_date = ship_date

    def run(self):
        try:
            self.stage.emit("正在解析 FBA 数据文件夹…")
            fba_result = parse_fba_folder(Path(self._fba_folder), self._ship_date)

            self.stage.emit("正在加载发货计划表…")
            wb = openpyxl.load_workbook(self._plan_path, data_only=False)
            if _MAIN_SHEET_NAME not in wb.sheetnames:
                self.failed.emit(f"发货计划表里没有找到「{_MAIN_SHEET_NAME}」这张表")
                return
            ws = wb[_MAIN_SHEET_NAME]

            self.stage.emit("正在校验匹配和数量…")
            plan: Plan = build_plan(ws, fba_result, self._ship_date)

            if plan.has_blocking_errors:
                self.blocked.emit(plan.errors)
                return

            self.stage.emit("正在拆分并写入…")
            apply_plan(ws, plan)

            self.stage.emit("正在备份并存盘…")
            backup_path = atomic_save_with_backup(wb, self._plan_path)

            report_path = ""
            if plan.skipped_fba_rows or plan.duplicates_removed:
                try:
                    written = write_skip_report(
                        plan.skipped_fba_rows, Path(self._fba_folder), plan.duplicates_removed
                    )
                    report_path = str(written) if written is not None else ""
                except Exception:
                    report_path = ""

            self.succeeded.emit(backup_path.name, len(plan.skipped_fba_rows), report_path)
        except Exception as exc:
            self.failed.emit(str(exc))


class FbaShipmentIdFillPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._worker: _WriteWorker | None = None
        self._settings = QSettings()

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        outer_layout.addWidget(scroll_area)

        content = QWidget()
        scroll_area.setWidget(content)
        layout = QVBoxLayout(content)

        title = QLabel("FBA 货件编号回填")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        note = QLabel(
            "把爬取到的 FBA 货件编号按标签（MSKU）匹配进发货计划表的「发货计划(2)」表，按箱数"
            "拆分对应的行——拆分出来的行会紧跟在原来那一行后面，不会挪到表格最下面。"
            "点「写入」直接生效，写入前会自动先备份原文件。如果这一批有数量对不上、目标行"
            "已经有 FBA ID 这类问题，会在写入前整批拦下来，列在下面，不会写进去一半。"
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        inputs_box = QGroupBox("输入")
        inputs_layout = QVBoxLayout(inputs_box)
        row, self._plan_edit = _file_picker_row("发货计划表", self._browse_plan)
        inputs_layout.addLayout(row)
        row, self._fba_folder_edit = _file_picker_row("FBA 数据文件夹", self._browse_fba_folder)
        inputs_layout.addLayout(row)

        date_row = QHBoxLayout()
        date_row.addWidget(QLabel("筛选发货日期（只比较月.日，忽略年份，去匹配货件名称里的日期）"))
        self._date_edit = QDateEdit()
        self._date_edit.setCalendarPopup(True)
        today = dt.date.today()
        self._date_edit.setDate(QDate(today.year, today.month, today.day))
        date_row.addWidget(self._date_edit)
        date_row.addStretch(1)
        inputs_layout.addLayout(date_row)
        layout.addWidget(inputs_box)

        self._plan_edit.setText(self._settings.value(_SETTINGS_KEY_PLAN, ""))
        self._fba_folder_edit.setText(self._settings.value(_SETTINGS_KEY_FBA_FOLDER, ""))

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
        self._error_text.setMaximumHeight(240)
        layout.addWidget(self._error_text)

    def _browse_plan(self):
        start_dir = str(Path(self._plan_edit.text()).parent) if self._plan_edit.text() else ""
        path, _ = QFileDialog.getOpenFileName(self, "选择发货计划表", start_dir, "Excel 文件 (*.xlsx)")
        if path:
            self._plan_edit.setText(path)
            self._settings.setValue(_SETTINGS_KEY_PLAN, path)

    def _browse_fba_folder(self):
        start_dir = self._fba_folder_edit.text() or ""
        path = QFileDialog.getExistingDirectory(self, "选择 FBA 数据文件夹", start_dir)
        if path:
            self._fba_folder_edit.setText(path)
            self._settings.setValue(_SETTINGS_KEY_FBA_FOLDER, path)

    def _start_write(self):
        plan_path = self._plan_edit.text().strip()
        fba_folder = self._fba_folder_edit.text().strip()

        missing = []
        if not plan_path:
            missing.append("发货计划表")
        if not fba_folder:
            missing.append("FBA 数据文件夹")
        if missing:
            QMessageBox.warning(self, "缺少输入", "还没选：" + "、".join(missing))
            return

        qdate = self._date_edit.date()
        ship_date = dt.date(qdate.year(), qdate.month(), qdate.day())

        self._write_button.setEnabled(False)
        self._progress.setRange(0, 0)
        self._progress.show()
        self._error_text.hide()
        self._status_label.setText("正在处理……表格行数多的话可能要一点时间，请耐心等待，不要关闭窗口。")

        self._worker = _WriteWorker(plan_path, fba_folder, ship_date)
        self._worker.blocked.connect(self._on_blocked)
        self._worker.succeeded.connect(self._on_succeeded)
        self._worker.failed.connect(self._on_failed)
        self._worker.stage.connect(self._on_stage)
        self._worker.start()

    def _on_stage(self, text: str):
        self._status_label.setText(text)

    def _on_blocked(self, error_lines: list):
        self._progress.hide()
        self._write_button.setEnabled(True)
        self._status_label.setText(f"这一批有 {len(error_lines)} 处问题，全部列在下面，一个都没写入。")
        self._error_text.setPlainText("\n".join(error_lines))
        self._error_text.show()

    def _on_succeeded(self, backup_name: str, skipped_count: int, report_path: str):
        self._progress.hide()
        text = f"已写入。备份文件：{backup_name}"
        if skipped_count > 0:
            text += f"\nFBA 数据里有 {skipped_count} 行被跳过（日期不匹配/已取消/数量异常）。"
            text += f"\n跳过说明表格：{report_path}" if report_path else "\n（跳过说明表格生成失败，只能自己核对）"
        self._status_label.setText(text)
        self._write_button.setText("已写入")

    def _on_failed(self, message: str):
        self._progress.hide()
        self._write_button.setEnabled(True)
        self._status_label.setText("写入失败，见弹窗说明。")
        QMessageBox.critical(self, "写入失败", message)

    def stop_running_tasks(self):
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(30000)
