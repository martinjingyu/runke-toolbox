"""海外仓销量汇总表导入——这个工具的界面。

流程是"生成预览 → 人工看一眼 → 确认写入"两步走（跟 logistics_tracking 那个工具一个模式，
不是 purchase_order_import 那种"点一下直接全部跑完"的模式）——因为这里改的是汇总表里
已经存在的销量数字，很容易跟"这一天到底导过没有"搞混，让人先看一眼"从多少改成多少"再确认
更安全。

真正落盘那一步不能用 openpyxl 存（会把目标表里 WPS 专有的单元格图片弄丢，见
xlsx_writer.py 开头的说明），走的是 planner.apply_plan -> xlsx_writer.apply_cell_updates
这条自己拼 zip 的路径，写之前还是会用 core.backup 里的机制先留一份备份。
"""
from __future__ import annotations

from pathlib import Path

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

from core.diff_preview import DiffPreviewGroup, DiffTable

from .planner import ImportPlan, PlanError, apply_plan, build_plan
from .platform_rules import IMPORTABLE_PLATFORMS, guess_platform_from_filename
from .source_import import SourceReadError, SourceRecord, read_castlegate_csv, read_erp_export

_SETTINGS_KEY_TARGET = "overseas_warehouse/sales_import/target_path"

_PLATFORM_LABELS = {
    "WF-RX": "Wayfair US（RX）",
    "WF-RQ": "Wayfair RQ",
    "WF-TS": "Wayfair3（TS）",
    "OS": "Overstock（OS）",
}
_LABEL_TO_PLATFORM = {v: k for k, v in _PLATFORM_LABELS.items()}


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


class _CsvFileEntry:
    def __init__(self, path: str):
        self.path = path
        self.platform_combo: QComboBox | None = None
        default_text = None
        guessed = guess_platform_from_filename(Path(path).name)
        if guessed:
            default_text = _PLATFORM_LABELS[guessed]
        self.rebuild_combo(default_text)

    @property
    def platform(self) -> str:
        return _LABEL_TO_PLATFORM[self.platform_combo.currentText()]

    def rebuild_combo(self, selected_text: str | None) -> None:
        """再次点「添加 CSV 文件…」时，_rebuild_csv_table 会先 setRowCount(0) 清空整张表——
        之前用 setCellWidget 挂上去的 QComboBox 归表格所有，这一清空就被 Qt 销毁了，
        这个 entry 手里存的 self.platform_combo 就变成了指向已经被删掉的 C++ 对象的
        Python 引用，再往表格里塞同一个对象会直接崩掉（不是抛异常能接住的那种，是
        C++ 层面的 use-after-free，表现出来就是整个程序闪退）。所以每次重建表格，
        不管是不是新加的文件，所有 entry 的下拉框都要重新建一份，不能指望旧的还能用——
        跟 shipment_plan_apply/panel.py 的 _PlanFileEntry.rebuild_combos 是同一个坑。
        """
        self.platform_combo = QComboBox()
        for key in IMPORTABLE_PLATFORMS:
            self.platform_combo.addItem(_PLATFORM_LABELS[key])
        if selected_text:
            self.platform_combo.setCurrentText(selected_text)


def _build_diff_table(plan: ImportPlan) -> DiffTable:
    headers = ["仓库", "平台", "SKU", "品名", "数量"]
    before_rows = []
    after_rows = []
    for d in plan.diff_rows:
        before_rows.append({
            "仓库": d.warehouse, "平台": d.platform, "SKU": d.sku,
            "品名": d.product_name, "数量": d.old_value,
        })
        after_rows.append({
            "仓库": d.warehouse, "平台": d.platform, "SKU": d.sku,
            "品名": d.product_name, "数量": d.new_value,
        })
    return DiffTable(headers=headers, before_rows=before_rows, after_rows=after_rows)


class _PreviewWorker(QThread):
    succeeded = Signal(object, list)  # ImportPlan, 源数据里被跳过的行（人工可读的说明）
    failed = Signal(str)
    stage = Signal(str)

    def __init__(self, target_path: str, day: int, erp_path: str, csv_files: list[tuple[str, str]]):
        super().__init__()
        self._target_path = target_path
        self._day = day
        self._erp_path = erp_path
        self._csv_files = csv_files

    def run(self):
        try:
            records: list[SourceRecord] = []
            source_skipped: list[str] = []
            if self._erp_path:
                self.stage.emit(f"正在读取 ERP 导出「{Path(self._erp_path).name}」…")
                result = read_erp_export(self._erp_path)
                records.extend(result.records)
                source_skipped.extend(result.skipped)

            for path, platform in self._csv_files:
                self.stage.emit(f"正在读取「{Path(path).name}」…")
                result = read_castlegate_csv(path, platform)
                records.extend(result.records)
                source_skipped.extend(result.skipped)

            if not records:
                self.failed.emit("没有读到任何源数据——ERP 导出和 CSV 至少要选一个")
                return

            self.stage.emit("正在跟目标表的表头对照、计算改动预览…")
            plan = build_plan(self._target_path, self._day, records)
        except (SourceReadError, PlanError) as exc:
            self.failed.emit(str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - 兜底，未预期的错误也要展示给人看而不是让线程静默死掉
            self.failed.emit(f"没预料到的错误：{exc}")
            return

        self.succeeded.emit(plan, source_skipped)


class _ApplyWorker(QThread):
    succeeded = Signal(object)  # 备份文件路径
    failed = Signal(str)

    def __init__(self, plan: ImportPlan):
        super().__init__()
        self._plan = plan

    def run(self):
        try:
            backup_path = apply_plan(self._plan)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))
            return
        self.succeeded.emit(backup_path)


class SalesImportPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._settings = QSettings()
        self._csv_entries: list[_CsvFileEntry] = []
        self._preview_worker: _PreviewWorker | None = None
        self._apply_worker: _ApplyWorker | None = None
        self._plan: ImportPlan | None = None

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        outer_layout.addWidget(scroll_area)

        content = QWidget()
        scroll_area.setWidget(content)
        layout = QVBoxLayout(content)

        title = QLabel("海外仓销量汇总表导入")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        note = QLabel(
            "把 ERP 导出 + CastleGate(CG) 平台 CSV 导出里的当天销量，写进《US库存销售明细表》"
            "对应仓库 sheet 的「几号」那一列。仓库按 ERP「发货仓库」字段自动判断"
            "（含IL→IL，含TX1→TX1，含CA1→CA1，其余都算CG）；CG 的 3 份 CSV 本身就是 CG 仓库的数据，"
            "平台按文件名默认猜一个，猜错了在下面表格里改。同一天同仓库同平台同 SKU 如果有多笔，"
            "数量会自动加总。先点「生成预览」看一眼改动，确认没问题再点「确认写入」，写入前会自动备份原文件。"
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        inputs_box = QGroupBox("输入")
        inputs_layout = QVBoxLayout(inputs_box)
        row, self._target_edit = _file_picker_row("US库存销售明细表（目标汇总表）", self._browse_target)
        inputs_layout.addLayout(row)

        day_row = QHBoxLayout()
        day_row.addWidget(QLabel("导入哪一天（几号）"))
        self._date_edit = QDateEdit()
        self._date_edit.setDisplayFormat("d号")
        self._date_edit.setDate(QDate.currentDate())
        day_row.addWidget(self._date_edit)
        day_row.addStretch(1)
        inputs_layout.addLayout(day_row)

        row, self._erp_edit = _file_picker_row("ERP 导出（可选，不选就跳过）", self._browse_erp)
        inputs_layout.addLayout(row)
        layout.addWidget(inputs_box)

        self._target_edit.setText(self._settings.value(_SETTINGS_KEY_TARGET, ""))

        csv_box = QGroupBox("CastleGate(CG) 平台 CSV 导出（可以一次选好几份）")
        csv_layout = QVBoxLayout(csv_box)
        add_row = QHBoxLayout()
        add_button = QPushButton("添加 CSV 文件…")
        add_button.clicked.connect(self._add_csv_files)
        add_row.addWidget(add_button)
        add_row.addStretch(1)
        csv_layout.addLayout(add_row)

        self._csv_table = QTableWidget(0, 3)
        self._csv_table.setHorizontalHeaderLabels(["文件", "平台", ""])
        self._csv_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        csv_layout.addWidget(self._csv_table)
        layout.addWidget(csv_box)

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

        self._unmatched_text = QTextEdit()
        self._unmatched_text.setReadOnly(True)
        self._unmatched_text.hide()
        self._unmatched_text.setMaximumHeight(120)
        layout.addWidget(self._unmatched_text)

        self._diff_group = DiffPreviewGroup("改动预览（仓库+平台+SKU）", key_fields=["仓库", "平台", "SKU"])
        layout.addWidget(self._diff_group)

    # ---- 输入 ----

    def _browse_target(self):
        start_dir = str(Path(self._target_edit.text()).parent) if self._target_edit.text() else ""
        path, _ = QFileDialog.getOpenFileName(self, "选择 US库存销售明细表", start_dir, "Excel 文件 (*.xlsx)")
        if path:
            self._target_edit.setText(path)
            self._settings.setValue(_SETTINGS_KEY_TARGET, path)

    def _browse_erp(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 ERP 导出文件", "", "Excel 文件 (*.xls *.xlsx);;所有文件 (*)"
        )
        if path:
            self._erp_edit.setText(path)

    def _add_csv_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择 CastleGate 平台 CSV 导出", "", "CSV 文件 (*.csv)")
        for path in paths:
            self._csv_entries.append(_CsvFileEntry(path))
        self._rebuild_csv_table()

    def _rebuild_csv_table(self) -> None:
        table = self._csv_table
        # setRowCount(0) 会把之前 setCellWidget 挂上去的下拉框一起销毁掉（表格拥有这些
        # 控件），所以清空之前先把每个 entry 当前选的平台记下来，清空之后给每个 entry
        # （不只是新加的）都重新建一份下拉框、把选择复原——见 _CsvFileEntry.rebuild_combo
        # 的说明，这是"一个一个添加 CSV 文件会直接闪退"这个问题的根因。
        selected_by_entry = {id(e): e.platform_combo.currentText() for e in self._csv_entries}
        table.setRowCount(0)
        for entry in self._csv_entries:
            entry.rebuild_combo(selected_by_entry[id(entry)])

            row = table.rowCount()
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(Path(entry.path).name))
            table.setCellWidget(row, 1, entry.platform_combo)
            remove_button = QPushButton("移除")
            remove_button.clicked.connect(lambda _checked, e=entry: self._remove_csv_entry(e))
            table.setCellWidget(row, 2, remove_button)
        table.resizeColumnsToContents()

    def _remove_csv_entry(self, entry: _CsvFileEntry) -> None:
        if entry in self._csv_entries:
            self._csv_entries.remove(entry)
        self._rebuild_csv_table()

    # ---- 预览 ----

    def _start_preview(self) -> None:
        target_path = self._target_edit.text().strip()
        erp_path = self._erp_edit.text().strip()

        if not target_path:
            QMessageBox.warning(self, "缺少输入", "还没选 US库存销售明细表")
            return
        if not erp_path and not self._csv_entries:
            QMessageBox.warning(self, "缺少输入", "ERP 导出和 CSV 至少要选一个")
            return

        day = self._date_edit.date().day()
        csv_files = [(e.path, e.platform) for e in self._csv_entries]

        self._preview_button.setEnabled(False)
        self._confirm_button.setEnabled(False)
        self._confirm_button.setText("确认写入")
        self._progress.setRange(0, 0)
        self._progress.show()
        self._unmatched_text.hide()
        self._diff_group.clear()
        self._status_label.setText("正在读取源数据、生成预览……")

        self._preview_worker = _PreviewWorker(target_path, day, erp_path, csv_files)
        self._preview_worker.succeeded.connect(self._on_preview_succeeded)
        self._preview_worker.failed.connect(self._on_preview_failed)
        self._preview_worker.stage.connect(self._on_stage)
        self._preview_worker.start()

    def _on_stage(self, text: str) -> None:
        self._status_label.setText(text)

    def _on_preview_succeeded(self, plan: ImportPlan, source_skipped: list[str]) -> None:
        self._progress.hide()
        self._preview_button.setEnabled(True)
        self._plan = plan

        lines = []
        if source_skipped:
            lines.append("【源数据里被跳过的行】")
            lines.extend(source_skipped)
        if plan.fuzzy_matches:
            if lines:
                lines.append("")
            lines.append("【SKU 对不上原文，去掉最后一段后缀之后匹配成功】")
            lines.extend(
                f"来源 {'、'.join(f.origins)} 里的 SKU「{f.source_sku}」（{f.quantity} 件），"
                f"匹配成了「{f.warehouse}」汇总表里的 SKU「{f.matched_sku}」，已经合并计入那一行"
                for f in plan.fuzzy_matches
            )
        if plan.unmatched:
            if lines:
                lines.append("")
            lines.append("【目标表里没能匹配上、已跳过】")
            lines.extend(
                f"「{u.warehouse}」{u.platform} 平台的 SKU「{u.sku}」（合计 {u.quantity} 件）：{u.reason}"
                for u in plan.unmatched
            )
        if lines:
            self._unmatched_text.setPlainText("\n".join(lines))
            self._unmatched_text.show()
        else:
            self._unmatched_text.hide()

        self._diff_group.fill(_build_diff_table(plan))

        self._status_label.setText(
            f"预览完成：{len(plan.diff_rows)} 个 SKU 有改动（其中 {len(plan.fuzzy_matches)} 个是靠去掉后缀"
            f"才匹配上的），{len(plan.unmatched)} 个 SKU 没匹配上，源数据里有 {len(source_skipped)} 行被跳过"
            "（见下方说明）。确认没问题的话点「确认写入」。"
        )
        self._confirm_button.setEnabled(bool(plan.diff_rows))

    def _on_preview_failed(self, message: str) -> None:
        self._progress.hide()
        self._preview_button.setEnabled(True)
        self._status_label.setText("预览失败，见弹窗说明。")
        QMessageBox.critical(self, "预览失败", message)

    # ---- 确认写入 ----

    def _confirm_write(self) -> None:
        if self._plan is None:
            return

        reply = QMessageBox.question(
            self,
            "确认写入",
            f"确定要把上面预览的改动写进：\n  {self._plan.target_path}\n\n写入前会先在原文件旁边自动生成一份备份。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self._preview_button.setEnabled(False)
        self._confirm_button.setEnabled(False)
        self._progress.setRange(0, 0)
        self._progress.show()
        self._status_label.setText("正在备份、写入……")

        self._apply_worker = _ApplyWorker(self._plan)
        self._apply_worker.succeeded.connect(self._on_apply_succeeded)
        self._apply_worker.failed.connect(self._on_apply_failed)
        self._apply_worker.start()

    def _on_apply_succeeded(self, backup_path) -> None:
        self._progress.hide()
        self._preview_button.setEnabled(True)
        self._status_label.setText(f"已写入。备份文件：{Path(backup_path).name}")
        self._confirm_button.setText("已写入")
        # 特意不重新启用——防止手滑再点一次把同样的改动重复写一遍。

    def _on_apply_failed(self, message: str) -> None:
        self._progress.hide()
        self._preview_button.setEnabled(True)
        self._confirm_button.setEnabled(True)
        self._status_label.setText("写入失败，见弹窗说明。")
        QMessageBox.critical(self, "写入失败", message)

    def stop_running_tasks(self) -> None:
        if self._preview_worker is not None and self._preview_worker.isRunning():
            self._preview_worker.wait(30000)
        if self._apply_worker is not None and self._apply_worker.isRunning():
            self._apply_worker.wait(30000)
