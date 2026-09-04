"""物流跟踪自动更新——这个工具的界面。

跟"发货计划自动更新"一样要小心：会真的改物流跟踪表格。所以也是"生成预览"只在内存里跑完
（联网查询 + 算出变化），人看过没问题再点"确认写入"，写入前自动备份原文件——预览组件直接复用
core/diff_preview.py，跟 shipment_plan_apply 是同一个模板组件。

查询本身要联网、要经过好几个物流商的账号，放在后台线程里跑（_TrackingPreviewWorker），
按"查完了几个物流商"报进度（见 tracking_pipeline.TrackingSheet.build_preview 的说明，查询
内部本身就是并行的，没法再拆更细的进度）。
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from PySide6.QtCore import QSettings, Qt, QThread, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
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
    QVBoxLayout,
    QWidget,
)

from core.backup import backup_file
from core.diff_preview import DiffPreviewGroup, DiffTable

from .credential_store import CredentialStore, StoredAccount
from .platforms.registry import PLATFORM_LABELS
from .tracking_pipeline import TrackingPreviewRow, TrackingSheet, list_sheet_names

_SETTINGS_KEY_PATH = "logistics_tracking/tracking_path"
_SHEET_NAME_HINT = "出货跟踪"


def _default_sheet_index(sheet_names: list[str]) -> int:
    for i, name in enumerate(sheet_names):
        if _SHEET_NAME_HINT in name:
            return i
    return 0


def _build_diff_table(previews: list[TrackingPreviewRow]) -> DiffTable:
    headers = ["行号", "物流商", "运单号", "货件状态", "最后流水", "是否有更新"]
    before_rows = []
    after_rows = []
    for p in previews:
        before_rows.append({
            "行号": p.row, "物流商": p.carrier, "运单号": p.waybill, "货件状态": p.status,
            "最后流水": p.old_last_route, "是否有更新": p.old_has_update,
        })
        after_rows.append({
            "行号": p.row, "物流商": p.carrier, "运单号": p.waybill, "货件状态": p.status,
            "最后流水": p.new_last_route, "是否有更新": p.new_has_update,
        })
    return DiffTable(headers=headers, before_rows=before_rows, after_rows=after_rows)


class _TrackingPreviewWorker(QThread):
    succeeded = Signal(object)  # (workbook, TrackingSheet, list[TrackingPreviewRow])
    failed = Signal(str)
    progress = Signal(int, int)  # 已查完的物流商数, 涉及的物流商总数

    def __init__(self, path: str, sheet_name: str, accounts_by_platform: dict):
        super().__init__()
        self._path = path
        self._sheet_name = sheet_name
        self._accounts_by_platform = accounts_by_platform

    def run(self):
        try:
            wb = openpyxl.load_workbook(self._path)
            sheet = TrackingSheet(wb[self._sheet_name])
            previews = sheet.build_preview(
                self._accounts_by_platform,
                progress_callback=lambda done, total: self.progress.emit(done, total),
            )
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.succeeded.emit((wb, sheet, previews))


class _CredentialDialog(QDialog):
    """账号管理：选一个货代平台（下拉框），下面列出这个平台已经配的账号（按查询优先顺序），
    可以上移/下移调整顺序、删除，也可以直接在下面的输入框填账号密码新增——改一次立刻存一次，
    不用另外点保存。
    """

    def __init__(self, store: CredentialStore, parent=None):
        super().__init__(parent)
        self.setWindowTitle("账号管理")
        self.resize(520, 420)
        self._store = store

        layout = QVBoxLayout(self)
        note = QLabel(
            "账号密码只存在这台电脑上，不会同步给别人、也不会跟着表格一起发出去。"
            "同一个货代平台可以配好几个账号，查询时按下面列表从上到下依次尝试，"
            "一个账号查到了就不用再试后面的。"
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        platform_row = QHBoxLayout()
        platform_row.addWidget(QLabel("货代平台"))
        self._platform_combo = QComboBox()
        for code, label in PLATFORM_LABELS.items():
            self._platform_combo.addItem(f"{label}（{code}）", code)
        self._platform_combo.currentIndexChanged.connect(self._reload_table)
        platform_row.addWidget(self._platform_combo, 1)
        layout.addLayout(platform_row)

        self._table = QTableWidget(0, 3)
        self._table.setHorizontalHeaderLabels(["账号", "密码", ""])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self._table)

        add_box = QGroupBox("添加账号")
        add_layout = QHBoxLayout(add_box)
        self._username_edit = QLineEdit()
        self._username_edit.setPlaceholderText("账号")
        add_layout.addWidget(self._username_edit)
        self._password_edit = QLineEdit()
        self._password_edit.setPlaceholderText("密码")
        self._password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        add_layout.addWidget(self._password_edit)
        show_cb = QCheckBox("显示")
        show_cb.toggled.connect(
            lambda checked: self._password_edit.setEchoMode(
                QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
            )
        )
        add_layout.addWidget(show_cb)
        add_button = QPushButton("添加")
        add_button.clicked.connect(self._add_account)
        add_layout.addWidget(add_button)
        layout.addWidget(add_box)

        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_button = QPushButton("关闭")
        close_button.clicked.connect(self.accept)
        close_row.addWidget(close_button)
        layout.addLayout(close_row)

        self._reload_table()

    def _current_platform_code(self) -> str:
        return self._platform_combo.currentData()

    def _reload_table(self) -> None:
        accounts = self._store.accounts_for(self._current_platform_code())
        self._table.setRowCount(len(accounts))
        for i, account in enumerate(accounts):
            self._table.setItem(i, 0, QTableWidgetItem(account.username))
            self._table.setItem(i, 1, QTableWidgetItem("•" * max(len(account.password), 4)))

            btns = QWidget()
            btns_layout = QHBoxLayout(btns)
            btns_layout.setContentsMargins(0, 0, 0, 0)
            up_btn = QPushButton("上移")
            up_btn.setEnabled(i > 0)
            up_btn.clicked.connect(lambda _checked, idx=i: self._move(idx, -1))
            btns_layout.addWidget(up_btn)
            down_btn = QPushButton("下移")
            down_btn.setEnabled(i < len(accounts) - 1)
            down_btn.clicked.connect(lambda _checked, idx=i: self._move(idx, 1))
            btns_layout.addWidget(down_btn)
            del_btn = QPushButton("删除")
            del_btn.clicked.connect(lambda _checked, idx=i: self._delete(idx))
            btns_layout.addWidget(del_btn)
            self._table.setCellWidget(i, 2, btns)
        self._table.resizeColumnsToContents()

    def _add_account(self) -> None:
        username = self._username_edit.text().strip()
        password = self._password_edit.text()
        if not username or not password:
            QMessageBox.warning(self, "缺少输入", "账号和密码都要填")
            return
        code = self._current_platform_code()
        accounts = self._store.accounts_for(code)
        accounts.append(StoredAccount(username=username, password=password))
        self._store.set_accounts(code, accounts)
        self._username_edit.clear()
        self._password_edit.clear()
        self._reload_table()

    def _move(self, idx: int, delta: int) -> None:
        code = self._current_platform_code()
        accounts = self._store.accounts_for(code)
        j = idx + delta
        if j < 0 or j >= len(accounts):
            return
        accounts[idx], accounts[j] = accounts[j], accounts[idx]
        self._store.set_accounts(code, accounts)
        self._reload_table()

    def _delete(self, idx: int) -> None:
        code = self._current_platform_code()
        accounts = self._store.accounts_for(code)
        accounts.pop(idx)
        self._store.set_accounts(code, accounts)
        self._reload_table()


class LogisticsTrackingPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._settings = QSettings()
        self._credential_store = CredentialStore(self._settings)
        self._worker: _TrackingPreviewWorker | None = None

        self._wb = None
        self._sheet = None
        self._previews: list[TrackingPreviewRow] = []
        self._path = ""

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        outer_layout.addWidget(scroll_area)

        content = QWidget()
        scroll_area.setWidget(content)
        layout = QVBoxLayout(content)

        title = QLabel("物流跟踪自动更新")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        note = QLabel(
            "会真的修改物流跟踪表格——请先点「生成预览」检查没问题，再点「确认写入」。"
            "确认写入前会自动备份原文件。只处理「货件状态」不含「完成」字样（或为空）的行。"
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        inputs_box = QGroupBox("输入")
        inputs_layout = QVBoxLayout(inputs_box)

        path_row = QHBoxLayout()
        path_row.addWidget(QLabel("物流跟踪表格"))
        self._path_edit = QLineEdit()
        self._path_edit.setReadOnly(True)
        path_row.addWidget(self._path_edit, 1)
        browse_button = QPushButton("浏览…")
        browse_button.clicked.connect(self._browse_path)
        path_row.addWidget(browse_button)
        inputs_layout.addLayout(path_row)

        sheet_row = QHBoxLayout()
        sheet_row.addWidget(QLabel("选哪个 sheet"))
        self._sheet_combo = QComboBox()
        sheet_row.addWidget(self._sheet_combo, 1)
        inputs_layout.addLayout(sheet_row)

        account_row = QHBoxLayout()
        account_button = QPushButton("账号管理…")
        account_button.clicked.connect(self._open_credential_dialog)
        account_row.addWidget(account_button)
        account_row.addStretch(1)
        inputs_layout.addLayout(account_row)

        layout.addWidget(inputs_box)

        saved_path = self._settings.value(_SETTINGS_KEY_PATH, "")
        if saved_path:
            self._set_path(saved_path)

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
        self._progress.hide()
        layout.addWidget(self._progress)

        self._status_label = QLabel("")
        self._status_label.setWordWrap(True)
        layout.addWidget(self._status_label)

        self._diff_group = DiffPreviewGroup(
            "物流跟踪表格 · 改动对比",
            key_fields=["行号", "物流商", "运单号", "货件状态"],
        )
        layout.addWidget(self._diff_group)

    # ---- 输入 ----

    def _set_path(self, path: str) -> None:
        self._path_edit.setText(path)
        self._sheet_combo.clear()
        try:
            sheet_names = list_sheet_names(path)
        except Exception as exc:
            QMessageBox.warning(self, "读取失败", f"「{Path(path).name}」读取失败：{exc}")
            return
        self._sheet_combo.addItems(sheet_names)
        if sheet_names:
            self._sheet_combo.setCurrentIndex(_default_sheet_index(sheet_names))

    def _browse_path(self) -> None:
        start_dir = str(Path(self._path_edit.text()).parent) if self._path_edit.text() else ""
        path, _ = QFileDialog.getOpenFileName(self, "选择物流跟踪表格", start_dir, "Excel 文件 (*.xlsx)")
        if path:
            self._set_path(path)
            self._settings.setValue(_SETTINGS_KEY_PATH, path)

    def _open_credential_dialog(self) -> None:
        dialog = _CredentialDialog(self._credential_store, self)
        dialog.exec()

    # ---- 预览 ----

    def _start_preview(self) -> None:
        path = self._path_edit.text().strip()
        sheet_name = self._sheet_combo.currentText()
        if not path or not sheet_name:
            QMessageBox.warning(self, "缺少输入", "请先选择物流跟踪表格和对应的 sheet")
            return

        accounts_by_platform = self._credential_store.accounts_by_platform()
        if not accounts_by_platform:
            reply = QMessageBox.question(
                self,
                "还没配账号",
                "还没在「账号管理」里配任何货代平台的账号密码，这次跑下来所有运单都会显示"
                "「还没配账号」。要不要先去配一下再跑？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                return

        self._path = path
        self._preview_button.setEnabled(False)
        self._confirm_button.setEnabled(False)
        self._progress.setRange(0, 0)
        self._progress.show()
        self._status_label.setText("正在扫描表格、并行查询各货代平台……")
        self._diff_group.clear()

        self._worker = _TrackingPreviewWorker(path, sheet_name, accounts_by_platform)
        self._worker.succeeded.connect(self._on_preview_succeeded)
        self._worker.failed.connect(self._on_preview_failed)
        self._worker.progress.connect(self._on_progress)
        self._worker.start()

    def _on_progress(self, done: int, total: int) -> None:
        if total <= 0:
            return
        if self._progress.maximum() != total:
            self._progress.setRange(0, total)
        self._progress.setValue(done)
        self._status_label.setText(f"正在查询：{done}/{total} 个货代平台完成……")

    def _on_preview_succeeded(self, payload) -> None:
        wb, sheet, previews = payload
        self._progress.hide()
        self._preview_button.setEnabled(True)
        self._wb = wb
        self._sheet = sheet
        self._previews = previews

        updated = sum(1 for p in previews if p.new_has_update == "有更新")
        failed = sum(1 for p in previews if p.new_last_route is None)
        self._status_label.setText(
            f"预览完成，共扫到 {len(previews)} 行需要处理，其中 {updated} 行有更新、"
            f"{failed} 行查询没拿到路由（原因见下表「是否有更新」列）。确认没问题的话点「确认写入」。"
        )
        self._diff_group.fill(_build_diff_table(previews))
        self._confirm_button.setEnabled(bool(previews))
        self._confirm_button.setText("确认写入")

    def _on_preview_failed(self, message: str) -> None:
        self._progress.hide()
        self._preview_button.setEnabled(True)
        self._status_label.setText("预览失败，见弹窗说明。")
        QMessageBox.critical(self, "预览失败", message)

    # ---- 确认写入 ----

    def _confirm_write(self) -> None:
        if self._wb is None or self._sheet is None:
            return

        reply = QMessageBox.question(
            self,
            "确认写入",
            f"确定要把上面预览的改动写进：\n  {self._path}\n\n写入前会先在原文件旁边自动生成一份备份。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            backup_path = backup_file(self._path)
        except Exception as exc:
            QMessageBox.critical(self, "备份失败", f"没能先备份原文件，写入已取消：{exc}")
            return

        try:
            self._sheet.apply_preview(self._previews)
            self._wb.save(self._path)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "写入失败",
                f"存盘失败：{exc}\n\n备份文件还在（{backup_path}），原文件可能已经部分改动，建议手动检查。",
            )
            return

        self._status_label.setText(f"已写入。备份文件：{backup_path.name}")
        self._confirm_button.setEnabled(False)
        self._confirm_button.setText("已写入")

    def stop_running_tasks(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(5000)
